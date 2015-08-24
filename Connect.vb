Imports gcornu.Interop
Imports gcornu.Interop.Extensibility
Imports Microsoft.Office.Core
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

<ComVisible(True), Guid("01234567-89ab-cdef-0123-456789abcdef"), ProgId("MyVBAAddin.Connect")> _
Public Class Connect
    Implements Extensibility.IDTExtensibility2

    Private _VBE As VBAExtensibility.VBE
    Private _AddIn As VBAExtensibility.AddIn

    ' Buttons created by the add-in
    Private WithEvents _myStandardCommandBarButton As CommandBarButton
    Private WithEvents _myToolsCommandBarButton As CommandBarButton
    Private WithEvents _myCodeWindowCommandBarButton As CommandBarButton
    Private WithEvents _myToolBarButton As CommandBarButton
    Private WithEvents _myCommandBarPopup1Button As CommandBarButton
    Private WithEvents _myCommandBarPopup2Button As CommandBarButton

    ' CommandBars created by the add-in
    Private _myToolbar As CommandBar
    Private _myCommandBarPopup1 As CommandBarPopup
    Private _myCommandBarPopup2 As CommandBarPopup

    Private Sub OnConnection(Application As Object, ConnectMode As Extensibility.ext_ConnectMode, _
       AddInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection

        Try

            _VBE = DirectCast(Application, VBAExtensibility.VBE)
            _AddIn = DirectCast(AddInInst, VBAExtensibility.AddIn)

            Select Case ConnectMode

                Case Extensibility.ext_ConnectMode.ext_cm_Startup
                    ' OnStartupComplete will be called

                Case Extensibility.ext_ConnectMode.ext_cm_AfterStartup
                    InitializeAddIn()

            End Select

        Catch ex As Exception

            MessageBox.Show(ex.ToString())

        End Try

    End Sub

    Private Sub OnDisconnection(RemoveMode As Extensibility.ext_DisconnectMode, _
       ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection

        Try

            Select Case RemoveMode

                Case ext_DisconnectMode.ext_dm_HostShutdown, ext_DisconnectMode.ext_dm_UserClosed

                    ' Delete buttons on built-in commandbars
                    If Not (_myStandardCommandBarButton Is Nothing) Then
                        _myStandardCommandBarButton.Delete()
                    End If

                    If Not (_myCodeWindowCommandBarButton Is Nothing) Then
                        _myCodeWindowCommandBarButton.Delete()
                    End If

                    If Not (_myToolsCommandBarButton Is Nothing) Then
                        _myToolsCommandBarButton.Delete()
                    End If

                    ' Disconnect event handlers
                    _myToolBarButton = Nothing
                    _myCommandBarPopup1Button = Nothing
                    _myCommandBarPopup2Button = Nothing

                    ' Delete commandbars created by the add-in
                    If Not (_myToolbar Is Nothing) Then
                        _myToolbar.Delete()
                    End If

                    If Not (_myCommandBarPopup1 Is Nothing) Then
                        _myCommandBarPopup1.Delete()
                    End If

                    If Not (_myCommandBarPopup2 Is Nothing) Then
                        _myCommandBarPopup2.Delete()
                    End If

            End Select

        Catch e As System.Exception
            System.Windows.Forms.MessageBox.Show(e.ToString)
        End Try
    End Sub

    Private Sub OnStartupComplete(ByRef custom As System.Array) _
       Implements IDTExtensibility2.OnStartupComplete

        InitializeAddIn()

    End Sub

    Private Sub OnAddInsUpdate(ByRef custom As System.Array) _
       Implements IDTExtensibility2.OnAddInsUpdate

    End Sub

    Private Sub OnBeginShutdown(ByRef custom As System.Array) _
       Implements IDTExtensibility2.OnBeginShutdown

    End Sub

    Private Function AddCommandBarButton(ByVal commandBar As CommandBar, caption As String) As CommandBarButton

        Dim commandBarButton As CommandBarButton
        Dim commandBarControl As CommandBarControl

        commandBarControl = commandBar.Controls.Add(MsoControlType.msoControlButton)
        commandBarButton = DirectCast(commandBarControl, CommandBarButton)

        commandBarButton.Caption = caption
        commandBarButton.FaceId = 59

        Return commandBarButton

    End Function

    Private Sub InitializeAddIn()

        ' Constants for names of built-in commandbars of the VBA editor
        Const TOOLS_COMMANDBAR_NAME As String = "Tools"

        ' Constants for names of commandbars created by the add-in
        Const MY_COMMANDBAR_POPUP1_NAME As String = "TemporaryCommandBarPopup1"

        ' Constants for captions of commandbars created by the add-in
        Const MY_COMMANDBAR_POPUP1_CAPTION As String = "Git For VBA"
        Const MY_TOOLBAR_CAPTION As String = "Git For VBA toolbar"

        ' Built-in commandbars of the VBA editor
        Dim toolsCommandBar As CommandBar

        ' Other variables
        Dim toolsCommandBarControl As CommandBarControl
        Dim position As Integer

        Try

            ' Retrieve some built-in commandbars
            toolsCommandBar = _VBE.CommandBars.Item(TOOLS_COMMANDBAR_NAME)

            ' ------------------------------------------------------------------------------------
            ' New toolbar
            ' ------------------------------------------------------------------------------------

            ' Add a new toolbar 
            _myToolbar = _VBE.CommandBars.Add(MY_TOOLBAR_CAPTION, MsoBarPosition.msoBarTop, System.Type.Missing, True)

            ' Add a new button on that toolbar
            _myToolBarButton = AddCommandBarButton(_myToolbar, "Button1")

            ' Make visible the toolbar
            _myToolbar.Visible = True

            ' ------------------------------------------------------------------------------------
            ' New submenu under the "Tools" menu
            ' ------------------------------------------------------------------------------------

            ' Add a new commandbar popup 
            _myCommandBarPopup1 = DirectCast(toolsCommandBar.Controls.Add( _
               MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, _
               toolsCommandBar.Controls.Count + 1, True), CommandBarPopup)

            ' Change some commandbar popup properties
            _myCommandBarPopup1.CommandBar.Name = MY_COMMANDBAR_POPUP1_NAME
            _myCommandBarPopup1.Caption = MY_COMMANDBAR_POPUP1_CAPTION

            ' Add a new button on that commandbar popup
            _myCommandBarPopup1Button = AddCommandBarButton(_myCommandBarPopup1.CommandBar, "Button2")

            ' Make visible the commandbar popup
            _myCommandBarPopup1.Visible = True

            ' Calculate the position of a new commandbar popup to the right of the "Tools" menu
            toolsCommandBarControl = DirectCast(toolsCommandBar.Parent, CommandBarControl)
            position = toolsCommandBarControl.Index + 1

        Catch e As System.Exception
            System.Windows.Forms.MessageBox.Show(e.ToString)
        End Try

    End Sub

    Private Sub _myToolBarButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, _
       ByRef CancelDefault As Boolean) Handles _myToolBarButton.Click

        MessageBox.Show("Clicked " & Ctrl.Caption)

    End Sub

    Private Sub _myToolsCommandBarButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, _
       ByRef CancelDefault As Boolean) Handles _myToolsCommandBarButton.Click

        MessageBox.Show("Clicked " & Ctrl.Caption)

    End Sub

    Private Sub _myStandardCommandBarButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, _
       ByRef CancelDefault As Boolean) Handles _myStandardCommandBarButton.Click

        MessageBox.Show("Clicked " & Ctrl.Caption)

    End Sub

    Private Sub _myCodeWindowCommandBarButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, _
       ByRef CancelDefault As Boolean) Handles _myCodeWindowCommandBarButton.Click

        MessageBox.Show("Clicked " & Ctrl.Caption)

    End Sub

    Private Sub _myCommandBarPopup1Button_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, _
       ByRef CancelDefault As Boolean) Handles _myCommandBarPopup1Button.Click

        MessageBox.Show("Clicked " & Ctrl.Caption)

    End Sub

End Class