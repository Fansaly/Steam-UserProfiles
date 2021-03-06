On Error Resume Next

Const HKEY_CLASSES_ROOT  = &H80000000
Const HKEY_CURRENT_USER  = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS         = &H80000003

Const ForReading         =  1  ' 只读模式
Const ForWriting         =  2  ' 只写模式
Const ForAppending       =  8  ' 文件末尾追加
Const TristateFalse      =  0  ' ASCII 格式
Const TristateTrue       = -1  ' Unicode 格式
Const TristateUseDefault = -2  ' 系统默认格式

Public strComputer
strComputer              = "." ' 本地计算机

Public OS: Set OS = GetOSInfos()

' 获取当前系统部分信息
Function GetOSInfos()
    Dim OS_: Set OS_ = CreateObject("Scripting.Dictionary")

    Dim ObjOS, O
    For Each ObjOS In GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
        For Each O In ObjOS.Properties_
            OS_.Add O.Name, O.Value
        Next
    Next

    Dim rEx_OSv, regEx_OSv
    rEx_OSv = "(\d+)\.(\d+).*"
    Set regEx_OSv = New RegExp
    regEx_OSv.Pattern = rEx_OSv
    regEx_OSv.IgnoreCase = False
    regEx_OSv.Global = True

    OS_.Add "NT", regEx_OSv.Replace(OS_.Item("Version"), "$1$2")

    Dim rEx_OSa, regEx_OSa
    rEx_OSa = "\d+"
    Set regEx_OSa = New RegExp
    regEx_OSa.Pattern = rEx_OSa
    regEx_OSa.IgnoreCase = False
    regEx_OSa.Global = True

    Dim Matches, Match, Architecture
    Set Matches = regEx_OSa.Execute(OS_.Item("OSArchitecture"))
    For Each Match In Matches
        Architecture = Architecture & Match.Value
    Next

    If (Architecture = 32) Then Architecture = 86

    OS_.Add "Architecture", "x" & Architecture

    Set regEx_OSv = Nothing
    Set regEx_OSa = Nothing

    Set GetOSInfos = OS_
End Function

Dim fso, WshShell, WshSysEnv, SystemRoot, FolderPath, TargetFolder
Set fso       = CreateObject("Scripting.FileSystemObject")
Set WshShell  = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("Process")
SystemRoot    = WshSysEnv.Item("SystemRoot")
FolderPath    = fso.GetParentFolderName(WScript.ScriptFullName)
TargetFolder  = FolderPath

Dim Parameters: Set Parameters = WScript.Arguments

Dim OutputFileName, SILENT, NOTICE
OutputFileName = Parameters(0)
SILENT =  Parameters(1)

If (UCase(SILENT) <> "SILENT") Then
    NOTICE = True
Else
    NOTICE = False
End If

If (NOTICE) Then WshShell.Popup "即将开始制作文件，请稍候...", 3, "提示信息", 0 + 64 + 4096

Dim temp_dir, tools_dir, assets_dir, config_dir, profiles_dir
temp_dir     = TargetFolder & "\temp"
tools_dir    = TargetFolder & "\Tools"
assets_dir   = TargetFolder & "\Assets"
config_dir   = TargetFolder & "\Config"
profiles_dir = TargetFolder & "\UserProfiles"

If Not fso.FolderExists(temp_dir) Then fso.CreateFolder(temp_dir)


' ResourceHacker 配置
Dim ResourceHacker, ResourceHacker_Parameter, ResourceHacker_Script, ResourceHacker_Log
ResourceHacker           = Chr(34)& tools_dir & "\ResourceHacker\ResourceHacker.exe" &Chr(34)
ResourceHacker_Parameter = " -script "
ResourceHacker_Script    = Chr(34)& temp_dir & "\resourcehacker_script.txt" &Chr(34)
ResourceHacker_Log       = Chr(34)& temp_dir & "\ResourceHacker.log" &Chr(34)

' 7zSFX 模块
Dim SFX: Set SFX = CreateObject("Scripting.Dictionary")
SFX.Add "x86", "7zsd_LZMA2.sfx"
SFX.Add "x64", "7zsd_LZMA2.sfx" ' 为避免 "Exception code: 0x000006ba", 本应为 7zsd_LZMA2_x64.sfx

Dim SFX_File, Cfg_7zSFX, Org_7zSFX, New_7zSFX
SFX_File = "7zSFXTools\" & SFX.Item(OS.Item("Architecture"))

Cfg_7zSFX   = Chr(34)& config_dir & "\config.txt" &Chr(34)
Org_7zSFX   = Chr(34)& tools_dir & "\" & SFX_File &Chr(34)
New_7zSFX   = Chr(34)& temp_dir & "\7zSFX.sfx" &Chr(34)


' 设置 icon 资源
Dim ICON: ICON = Chr(34)& TargetFolder & "\Assets\images\SteamPony.ico" &Chr(34)


' 修改图标
' ResourceHacker Script
Dim Stream: Set Stream = CreateObject("Adodb.Stream")
Stream.Type = 2
Stream.Mode = 3
Stream.Charset = "UTF-8"
Stream.Open
Stream.WriteText "[FILENAMES]", 1
Stream.WriteText "Exe = " & Org_7zSFX, 1
Stream.WriteText "SaveAs = " & New_7zSFX, 1
Stream.WriteText "Log = " & ResourceHacker_Log, 1
Stream.WriteText "", 1
Stream.WriteText "[COMMANDS]", 1
Stream.WriteText "-modify " & ICON & ", icongroup,101,1049", 1
Stream.SaveToFile Replace(ResourceHacker_Script, Chr(34), ""), 2
Stream.Close
Set Stream = Nothing

' ResourceHacker command
WshShell.Run ResourceHacker & ResourceHacker_Parameter & ResourceHacker_Script, 0, True


' 7-Zip 配置
Dim Execute_7zip, Execute_Parameter0, Execute_Parameter1, Execute_Parameter2, Execute_OutputFileName, Execute_OutputArchive, Execute_PackageFiles, TheProgram
Execute_7zip            = Chr(34)& tools_dir & "\7-Zip\" & OS.Item("Architecture") & "\7z.exe" &Chr(34)
' 参数
Execute_Parameter0      = "a -t7z"

' 文件名
If (OutputFileName <> "") Then
    Execute_OutputFileName  = OutputFileName
Else
    Execute_OutputFileName  = "Steam_UserProfiles"
End If

' 7z 文件位置
Execute_OutputArchive   = Chr(34)& TargetFolder & "\" & Execute_OutputFileName & ".7z" &Chr(34)
' 将要打包的文件
Execute_PackageFiles    = " " & Chr(34)& tools_dir    & "\" &Chr(34)&_
                          " " & Chr(34)& assets_dir   & "\" &Chr(34)&_
                          " " & Chr(34)& config_dir   & "\" &Chr(34)&_
                          " " & Chr(34)& profiles_dir & "\" &Chr(34)&_
                          " " & Chr(34)& TargetFolder & "\make.vbs" &Chr(34)&_
                          " " & Chr(34)& TargetFolder & "\SteamUserProfiles.hta" &Chr(34)
' 参数
Execute_Parameter1      = "-xr!ResourceHacker.ini"
' 参数
Execute_Parameter2      = "-mx=9 -ms -mf -mhc -mmt -m0=LZMA2:a=1:d=26:mf=bt4:fb=64"
' exe 文件位置
TheProgram              = Chr(34)& TargetFolder & "\" & Execute_OutputFileName & ".exe" &Chr(34)

If fso.FileExists( Replace(Execute_OutputArchive, Chr(34), "") ) Then fso.DeleteFile( Replace(Execute_OutputArchive, Chr(34), "") )

' 打包文件
WshShell.Run Execute_7zip           & " " &_
             Execute_Parameter0     & " " &_
             Execute_OutputArchive  & " " &_
             Execute_PackageFiles   & " " &_
             Execute_Parameter1     & " " &_
             Execute_Parameter2, 0, True

' 制作 exe 文件
WshShell.Run "cmd.exe /C copy /y /b " &_
             New_7zSFX & " + " &_
             Cfg_7zSFX & " + " &_
             Execute_OutputArchive & " " &_
             TheProgram, 0, True

If (NOTICE) Then WshShell.Popup "制作完成，请在当前目录查看。", 7, "提示信息", 0 + 64 + 4096

fso.DeleteFile( Replace(Execute_OutputArchive, Chr(34), "") )
