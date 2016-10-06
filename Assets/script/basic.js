masterFile = location.pathname; // 获取当前文件完整路径
parentFolder = masterFile.substring(0, masterFile.lastIndexOf("\\")); // 获取当前路径
// alert(parentFolder + '\n' + masterFile);

     fso = new ActiveXObject("Scripting.FileSystemObject");
   Shell = new ActiveXObject("Shell.Application");
WshShell = new ActiveXObject("WScript.Shell");

ForReading          =  1; // 以只读模式打开文件
ForWriting          =  2; // 以只写方式打开文件
ForAppending        =  8; // 打开文件并在文件末尾进行写操作
TristateFalse       =  0; // 以 ASCII 格式打开文件
TristateTrue        = -1; // 以 Unicode 格式打开文件
TristateUseDefault  = -2; // 以系统默认格式打开文件

// 获取 Windows 系统信息
// https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx
getOperatingSystem = function() {
    var WbemLocator = new ActiveXObject("WbemScripting.SWbemLocator");
    var objWMIService = WbemLocator.ConnectServer(strComputer, "root\\cimv2");
    var colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem");

    var os = new Object();
    var enumItems = new Enumerator(colItems);
    for (; !enumItems.atEnd(); enumItems.moveNext()) {
        var properties = new Enumerator(enumItems.item().Properties_);
        for (; !properties.atEnd(); properties.moveNext()) {
            var name = properties.item().name;
            var value = properties.item().value ? properties.item().value : null;
            os[name] = value;
        }
    }

    if (isEmpty(os)) {
        return false;
    } else {
        var Architecture  = parseInt(os.OSArchitecture.match(/\d+/g));
        Architecture = Architecture == 32 ? 86 : Architecture;

        os["Architecture"]  = "x" + Architecture;
        os["OSVersion"]     = parseInt(os.Version.replace(/(\d+)\.(\d+).*/g, "$1$2"));

        return os;
    }

}
OS = getOperatingSystem();

// 如果是 vista 或以上版本系统以“管理员身份运行”
// if (OS.OSVersion >= 60) {
//     var file = masterFile;
//     // 通过扩展名的大小写，来区分第几次运行 hta
//     if (file.substring(file.length-4) != ".HTA") {
//         // 以 runas 方式重新启动该 hta
//         Shell.ShellExecute("mshta.exe", "\"" + file.substring(0, file.length-4) + ".HTA\"", "", "runas", 1);
//         // 退出当前 hta
//         window.close();
//         exit(0);
//     }
// }

// 获取任务栏的位置
getTaskbarLocation = function() {
    var strKeyPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StuckRects2"; // Windows 7
    var strValueName = "Settings";
    var arrBinaryValue;

    arrBinaryValue = getRegBinaryValue(HKEY_CURRENT_USER, strKeyPath, strValueName);

    if (!arrBinaryValue) {
        strKeyPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StuckRects3"; // Windows 10
        arrBinaryValue = getRegBinaryValue(HKEY_CURRENT_USER, strKeyPath, strValueName);
    }

    // alert(arrBinaryValue);

    return arrBinaryValue[12]; // 0-左 1-上 2-右 3-下
}

// 应用程序窗口位置
dialogSize = function(w, h) {
    var screenWidth = window.screen.width;
    var screenHeight = window.screen.height;
    var workspaceWidth = window.screen.availWidth;
    var workspaceHeight = window.screen.availHeight;

    var taskbarLocation = getTaskbarLocation();

    var left = (workspaceWidth - w) / 2;
    var top = (workspaceHeight - h) / 2;

    if (taskbarLocation == 0) {
        left += screenWidth - workspaceWidth;
    } else if (taskbarLocation == 1) {
        top += screenHeight - workspaceHeight;
    }

    while (true) {
        try {
            window.resizeTo(w, h);
            window.moveTo(left, top);
            break;
        } catch (e) { continue; }
    }

    return true;
}
dialogSize(600, 328);

// 退出应用程序执行的操作
quitApp = function() {
    self.close();
    // DeleteFolder(parentFolder + "\\Toolkit");
    // fso.DeleteFile(masterFile);
}
// quitApp();