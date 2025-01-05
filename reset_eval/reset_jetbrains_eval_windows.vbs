' reset jetbrains ide evals v1.0.4

' 判断是否存在idea64.exe进程，如果存在则提示用户关闭IDE，并退出脚本
Set oShell = CreateObject("WScript.Shell")
Set oProc = oShell.Exec("tasklist /FI ""IMAGENAME eq idea64.exe""")
If InStr(oProc.StdOut.ReadAll, "idea64.exe") > 0 Then
    MsgBox "Please close JetBrains IDE first.", vbExclamation, "Warning"
    WScript.Quit 1
End If

' 创建 WScript.Shell 对象，用于执行系统命令和访问环境变量
Set oShell = CreateObject("WScript.Shell")

' 创建 FileSystemObject 对象，用于文件和目录操作
Set oFS = CreateObject("Scripting.FileSystemObject")

' 获取用户主目录路径（如 C:\Users\username）
sHomeFolder = oShell.ExpandEnvironmentStrings("%USERPROFILE%")

' 获取 JetBrains IDE 数据目录路径（如 C:\Users\username\AppData\Roaming\JetBrains）
sJBDataFolder = oShell.ExpandEnvironmentStrings("%APPDATA%") + "\JetBrains"

' 创建正则表达式对象，用于匹配 JetBrains IDE 文件夹名称
Set re = New RegExp
re.Global     = True  ' 匹配所有符合条件的子字符串
re.IgnoreCase = True  ' 忽略大小写
re.Pattern    = "\.?(IntelliJIdea|GoLand|CLion|PyCharm|DataGrip|RubyMine|AppCode|PhpStorm|WebStorm|Rider).*"  ' 匹配 JetBrains IDE 文件夹名称

' 定义移除评估信息的子程序
Sub removeEval(ByVal file, ByVal sEvalPath)
    ' 检查文件夹名称是否符合 JetBrains IDE 文件夹名称模式
    bMatch = re.Test(file.Name)
    If Not bMatch Then
        Exit Sub  ' 如果不符合，则退出子程序
    End If

    ' 如果存在评估信息文件夹，则删除该文件夹
    If oFS.FolderExists(sEvalPath) Then
        oFS.DeleteFolder sEvalPath, True
    End If

    ' 修改 other.xml 文件，移除与评估信息相关的行
    content = ""
    otherFile = oFS.GetParentFolderName(sEvalPath) + "\options\other.xml"
    If oFS.FileExists(otherFile) Then
        Set txtStream = oFS.OpenTextFile(otherFile, 1, False)  ' 打开文件以读取模式
        Do While Not txtStream.AtEndOfStream
            line = txtStream.ReadLine
            ' 如果行中不包含 "name=""evlsprt""，则保留该行
            If InStr(line, "name=""evlsprt") = 0 Then
                content = content + line + vbLf
            End If
        Loop
        txtStream.Close

        ' 重新打开文件以写入模式，覆盖原文件内容
        Set txtStream = oFS.OpenTextFile(otherFile, 2, False)
        txtStream.Write content
        txtStream.Close
    End If
End Sub

' 遍历用户主目录下的所有子文件夹，调用 removeEval 子程序处理每个文件夹
If oFS.FolderExists(sHomeFolder) Then
    For Each oFile In oFS.GetFolder(sHomeFolder).SubFolders
        removeEval oFile, sHomeFolder + "\" + oFile.Name + "\config\eval"
    Next
End If

' 遍历 JetBrains 数据目录下的所有子文件夹，调用 removeEval 子程序处理每个文件夹
If oFS.FolderExists(sJBDataFolder) Then
    For Each oFile In oFS.GetFolder(sJBDataFolder).SubFolders
        removeEval oFile, sJBDataFolder + "\" + oFile.Name + "\eval"
    Next
End If

' 删除注册表中与 JetBrains 用户 ID 和设备 ID 相关的键值
On Error Resume Next  ' 忽略错误继续执行
oShell.RegDelete "HKEY_CURRENT_USER\Software\JavaSoft\Prefs\/Jet/Brains./User/Id/On/Machine"
oShell.RegDelete "HKEY_CURRENT_USER\Software\JavaSoft\Prefs\jetbrains\device_id"
oShell.RegDelete "HKEY_CURRENT_USER\Software\JavaSoft\Prefs\jetbrains\user_id_on_machine"

' 删除 JetBrains 数据目录下的特定文件
oFs.DeleteFile sJBDataFolder + "\bl"
oFs.DeleteFile sJBDataFolder + "\crl"
oFs.DeleteFile sJBDataFolder + "\PermanentUserId"
oFs.DeleteFile sJBDataFolder + "\PermanentDeviceId"

' 显示完成消息框
MsgBox "done"