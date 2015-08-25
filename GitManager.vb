Imports System.Windows.Forms

Public Class GitManager

    Dim _wsh As Object

    Public Sub New()
        _wsh = CreateObject("WScript.Shell")
    End Sub

    Protected Overrides Sub Finalize()

        _wsh = Nothing

    End Sub

    ' Inspired from https://github.com/JohnGreenan/5_ExcelVBE/blob/master/ExcelGit.bas
    Public Sub PerformGitActions()
        'Const
        Const strSourceDirectory As String = "D:\Users\Gauthier\Downloads\TestVBAGit"
        Const strCMD As String = "cmd /K"
        Const strChangeDirectoryTo As String = "cd"
        Const strGitAdd As String = "git add ."
        Const strGitCommit As String = "git commit -am"
        Const strGitPush As String = "git push"
        Const strGitStatus As String = "git status"
        Const strProcessID As String = "PID="
        Const strTitle As String = "Git Integration"
        'Variables
        Dim dtNow As Date
        Dim strTextFromStdStream As String
        Dim strBuiltCommand As String
        Dim strUserName As String
        Dim commitMessage As String

        'mWshell = New IWshRuntimeLibrary.WshShell
        'mWsh = New IWshRuntimeLibrary.WshNetwork

        dtNow = Now()

        'commitMessage = InputBox("Commit message:")
        commitMessage = "test"

        'Call ExportVBAFiles()

        '   Change to the correct folder with cmd:>cd folder
        strBuiltCommand = strCMD

        With _wsh.Exec(strBuiltCommand)

            ' Change directory
            strBuiltCommand = strChangeDirectoryTo & " " & strSourceDirectory
            .StdIn.WriteLine(strBuiltCommand)

            ' Track files (git add .)
            strBuiltCommand = strGitAdd
            .StdIn.WriteLine(strBuiltCommand)

            ' Commit files (git commit -am)
            strBuiltCommand = strGitCommit & " " & """" & dtNow & ":" & " " & commitMessage & """"
            .StdIn.WriteLine(strBuiltCommand)

            ' Push commit (git push)
            strBuiltCommand = strGitPush
            .StdIn.WriteLine(strBuiltCommand)

            'Cleanup
            .StdIn.Close()

            Do While Not .StdOut.AtEndOfStream
                strTextFromStdStream = "[" & strProcessID & .ProcessID & "]" & .StdOut.ReadLine()
                'Debug.Print(strTextFromStdStream)
                MessageBox.Show(strTextFromStdStream)
            Loop

            Do While Not .StdErr.AtEndOfStream
                strTextFromStdStream = "[" & strProcessID & .ProcessID & "]" & .StdErr.ReadLine()
                'Debug.Print(strTextFromStdStream)
                MessageBox.Show(strTextFromStdStream)
            Loop

            .StdErr.Close()
            .StdOut.Close()
            '.Terminate()
        End With

    End Sub
End Class
