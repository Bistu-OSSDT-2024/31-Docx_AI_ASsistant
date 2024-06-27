Imports System.Diagnostics
Imports System.IO
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles 启动纠错或总结.Click
        ' 获取当前插件的安装路径
        Dim pluginPath As String = AppDomain.CurrentDomain.BaseDirectory
        'MsgBox(pluginPath)
        ' 设置可执行文件相对路径
        Dim exeRelativePath As String = "main.exe"
        Dim exeFullPath As String = Path.Combine(pluginPath, exeRelativePath)

        ' 创建进程启动信息
        Dim startInfo As New ProcessStartInfo()
        startInfo.FileName = exeFullPath ' 设置可执行文件路径

        Try
            ' 启动可执行文件进程
            Using process As Process = Process.Start(startInfo)
                process.WaitForExit() ' 等待进程执行完毕
                MsgBox("执行完毕！")
            End Using
        Catch ex As Exception
            ' 处理启动进程时的异常
            MsgBox("Error: " & ex.Message)
        End Try
    End Sub
End Class
