
/**
 * 检查一个“对象”是否为空
 */
isEmpty = function(obj) {
    for (var key in obj) {
        if (obj.hasOwnProperty(key)) {
            return false;
        }
    }
    return true;
}


/**
 * 删除指定的文件
 */
DeleteFile = function(filespec) {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (fso.FileExists(filespec)) fso.DeleteFile(filespec);
}

/**
 * 删除指定的文件夹和其中的内容
 */
DeleteFolder = function(filespec) {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (fso.FolderExists(filespec)) fso.DeleteFolder(filespec);
}
/**
 * 创建指定的文件夹
 */
CreateFolder = function(filespec) {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (!fso.FolderExists(filespec)) fso.CreateFolder(filespec);
}


/**
 * 设置注册表值: string
 */
setRegStringValue = function(RootKey, SubKeyName, ValueName, Value) {
    var WbemLocator = new ActiveXObject("WbemScripting.SWbemLocator");
    var objWMIService = WbemLocator.ConnectServer(strComputer, "root\\default");
    var objReg = objWMIService.Get("StdRegProv");

    objMethod = objReg.Methods_.Item("SetStringValue");
    objInParam = objMethod.InParameters.SpawnInstance_();

    objInParam.hDefKey = RootKey;
    objInParam.sSubKeyName = SubKeyName;
    objInParam.sValueName = ValueName;
    objInParam.sValue = Value;

    var objOutParam = objReg.ExecMethod_("SetStringValue", objInParam);

    return objOutParam;
}

/**
 * 读取注册表值: string
 */
getRegStringValue = function(RootKey, SubKeyName, ValueName) {
    var WbemLocator = new ActiveXObject("WbemScripting.SWbemLocator");
    var objWMIService = WbemLocator.ConnectServer(strComputer, "root\\default");
    var objReg = objWMIService.Get("StdRegProv");

    objMethod = objReg.Methods_.Item("GetStringValue");
    objInParam = objMethod.InParameters.SpawnInstance_();

    objInParam.hDefKey = RootKey;
    objInParam.sSubKeyName = SubKeyName;
    objInParam.sValueName = ValueName;

    var objOutParam = objReg.ExecMethod_("GetStringValue", objInParam);

    if (objOutParam.sValue != null) {
        return objOutParam.sValue;
    } else {
        return false;
    }
}


/**
 * 读取注册表值: binary
 */
getRegBinaryValue = function(RootKey, SubKeyName, ValueName) {
    var WbemLocator = new ActiveXObject("WbemScripting.SWbemLocator");
    var objWMIService = WbemLocator.ConnectServer(strComputer, "root\\default");
    var objReg = objWMIService.Get("StdRegProv");

    objMethod = objReg.Methods_.Item("GetBinaryValue");
    objInParam = objMethod.InParameters.SpawnInstance_();

    objInParam.hDefKey = RootKey;
    objInParam.sSubKeyName = SubKeyName;
    objInParam.sValueName = ValueName;

    var objOutParam = objReg.ExecMethod_("GetBinaryValue", objInParam);

    if (objOutParam.uValue != null) {
        return new VBArray( objOutParam.uValue ).toArray();
    } else {
        return false;
    }
}


/**
 * 设置 x86, x64 正确的注册表节点
 */
setRightRegNodePath = function(strKeyPath, OSArchitecture) {
    var nodeName = {
        "x86" : "",
        "x64" : "WOW6432Node"
    };
    var strNode = nodeName[OSArchitecture]
        ? "\\" + nodeName[OSArchitecture]
        : nodeName[OSArchitecture];

    var re = /WOW\d{4,}Node/i;

    return strKeyPath.replace(/([A-Za-z]+)\\(\w+)(.*)/, function($0,$1,$2,$3){
        return $1 +
            (
                $2.replace(re, '')
                ? strNode + "\\" + $2
                : strNode
            ) +
            $3;
    });
}


/**
 * 提示手动退出一个程序
 */
exitProgram = function(appName) {
    var WbemLocator = new ActiveXObject("WbemScripting.SWbemLocator");
    var objWMIService = WbemLocator.ConnectServer(strComputer, "root\\cimv2");

    var app = new Object();

    do {
        app = {};

        var colItems = objWMIService.ExecQuery("Select * from Win32_Process where Name='" + appName + "'");

        var enumItems = new Enumerator(colItems);

        for (; !enumItems.atEnd(); enumItems.moveNext()) {
            var properties = new Enumerator(enumItems.item().Properties_);

            for (; !properties.atEnd(); properties.moveNext()) {
                var name = properties.item().name;
                var value = properties.item().value ? properties.item().value : null;
                app[name] = value;
            }
        }

        if (!isEmpty(app)) vbMsgBox('请退出 ' + appName + ' 后继续 ...', 48, '提示信息');

    } while (!isEmpty(app));
}


/**
 * 获取运行中程序的路径
 */
getProcessAppProperties = function(appName) {
    var WbemLocator = new ActiveXObject("WbemScripting.SWbemLocator");
    var objWMIService = WbemLocator.ConnectServer(strComputer, "root\\cimv2");
    var colItems = objWMIService.ExecQuery("Select * from Win32_Process where Name='" + appName + "'");

    var app = new Object();
    var enumItems = new Enumerator(colItems);

    for (; !enumItems.atEnd(); enumItems.moveNext()) {
        var properties = new Enumerator(enumItems.item().Properties_);

        for (; !properties.atEnd(); properties.moveNext()) {
            var name = properties.item().name;
            var value = properties.item().value ? properties.item().value : null;
            app[name] = value;
        }
    }

    return isEmpty(app) ? false : formatPath(app.ExecutablePath);
}


/**
 * 路径状态信息
 */
checkWindowsPathTips = function(str) {
    var obj = eval('(' + str.replace(/\\/g, '\\\\') + ')');
    var msg = obj.string;

    if (obj.status == 300) {
        msg = '正确：' + msg;
    } else if (obj.status == 200) {
        msg = '不能包含这些字符：' + msg;
    } else if (obj.status == 100) {
        msg = '磁盘驱动器错误：' + msg;
    } else {
        msg = '未知错误：' + msg;
    }

    return msg;
}

/**
 * 检测路径是否符合一定规范（非严格的）
 */
checkWindowsPath = function(str) {
    var re_h = /((.*)(?:\:)\\?)(.*)/,

        // Windows 默认文件名不能包括的字符 \/:*?"<>|
        re_b = /[^`~!@#\$%\^&\(\)-\+=\[\]\{\};',\.\\\w\s]/g,

        re_s = /[A-Za-z]/;

    return re_h.test(str)
        ? str.replace(re_h, function($0, $1, $2, $3){
            return $2.length == 1 && re_s.test($2)
                ? $3.match(re_b)
                    ? '{"status": 200, "string": "' + $3.match(re_b) + '"}'
                    : '{"status": 300, "string": "' + $1 + $3 + '"}'
                : '{"status": 100, "string": "' + $1 + '"}';
        })
        : '{"status": 0, "string": "' + str + '"}';
}


/**
 * 格式化路径
 */
formatPath = function(path, backslash) {
    path = path.substring(0, path.lastIndexOf("\\"));

    return backslash ? path + "\\" : path;
}


/**
 * 获取本地时间
 */
getLocalTime = function() {
    var D = new Date(), date = {};

    date.y = D.getFullYear();
    date.m = D.getMonth() + 1;
    date.d = D.getDate();
    date.H = D.getHours();
    date.i = D.getMinutes();
    date.s = D.getSeconds();

    date.m = date.m < 10 ? date.m = '0' + date.m : date.m;
    date.d = date.d < 10 ? date.d = '0' + date.d : date.d;
    date.H = date.H < 10 ? date.H = '0' + date.H : date.H;
    date.i = date.i < 10 ? date.i = '0' + date.i : date.i;
    date.s = date.s < 10 ? date.s = '0' + date.s : date.s;

    var time = date.y + "-" +
               date.m + "-" +
               date.d + ' ' +
               date.H + ":" +
               date.i + ":" +
               date.s;

    return time;
}
