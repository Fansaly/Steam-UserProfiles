/**
 * “对象”是否为空
 */
function isEmpty(obj) {
  for (var key in obj) {
    if (obj.hasOwnProperty(key)) {
      return false;
    }
  }
  return true;
}


/**
 * 获取 Windows 系统信息
 * https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx
 */
function getOperatingSystemInfo() {
  var objWMIService = WbemLocator.ConnectServer(strComputer, 'root\\cimv2');
  var colItems = objWMIService.ExecQuery('Select * from Win32_OperatingSystem');

  var os = new Object();
  var enumItems = new Enumerator(colItems);
  for (; !enumItems.atEnd(); enumItems.moveNext()) {
    var properties = new Enumerator(enumItems.item().Properties_);
    for (; !properties.atEnd(); properties.moveNext()) {
      var name = properties.item().name;
      var value = properties.item().value;
      os[name] = value ? value : null;
    }
  }

  if (isEmpty(os)) {
    return false;
  } else {
    var Architecture  = parseInt(os.OSArchitecture.match(/\d+/g));
    Architecture = Architecture == 32 ? 86 : Architecture;

    os['Architecture']  = 'x' + Architecture;
    os['NT']            = parseInt(os.Version.replace(/(\d+)\.(\d+).*/g, '$1$2'));

    return os;
  }
}


/**
 * 删除指定的文件
 */
function deleteFile(file) {
  if (fso.FileExists(file)) {
    fso.DeleteFile(file);
  }
}
/**
 * 删除指定的文件夹和其中的内容
 */
function deleteFolder(folder) {
  if (fso.FolderExists(folder)) {
    fso.DeleteFolder(folder);
  }
}
/**
 * 创建指定的文件夹
 */
function createFolder(folder) {
  if (!fso.FolderExists(folder)) {
    fso.CreateFolder(folder);
  }
}

/**
 * 复制文件到桌面
 * https://msdn.microsoft.com/en-us/library/windows/desktop/bb787866(v=vs.85).aspx
 * https://msdn.microsoft.com/en-us/library/windows/desktop/ms723207(v=vs.85).aspx
 */
function copyToDesktop(file) {
  if (!fso.FileExists(file)) {
    return;
  }

  var desktopDir = WshShell.SpecialFolders('Desktop');
  var objFolder = new Object;
  var result = false;

  try {
    objFolder = AppShell.NameSpace(desktopDir);

    if (objFolder != null) {
      objFolder.CopyHere(file, 8);
      result = true;
    }
  } catch (e) {}

  return result;
}


/**
 * 读取文件内容
 */
function readNormalContent(file) {
  var content = '';

  if (fso.FileExists(file)) {
    f = fso.OpenTextFile(file, ForReading, true, TristateFalse);
    content = f.ReadAll();
    f.Close();
  }

  return content;
}

/**
 * 读取 UTF-8 文件内容
 */
function readUTF8Content(file) {
  var content = '';

  if (fso.FileExists(file)) {
    var stream = new ActiveXObject('ADODB.Stream');

    stream.Type = 2;
    stream.Mode = 3;
    stream.Charset = 'UTF-8';
    stream.Open;
    stream.LoadFromFile(file);
    content = stream.ReadText;
    stream.Close;
  }

  return content;
}


/**
 * 设置注册表值: string
 */
function setRegStringValue(RootKey, SubKeyName, ValueName, Value) {
  var objWMIService = WbemLocator.ConnectServer(strComputer, 'root\\default');
  var objReg = objWMIService.Get('StdRegProv');

  var objMethod = objReg.Methods_.Item('SetStringValue');
  var objInParam = objMethod.InParameters.SpawnInstance_();

  objInParam.hDefKey = RootKey;
  objInParam.sSubKeyName = SubKeyName;
  objInParam.sValueName = ValueName;
  objInParam.sValue = Value;

  var objOutParam = objReg.ExecMethod_('SetStringValue', objInParam);

  return objOutParam;
}

/**
 * 读取注册表值: string
 */
function getRegStringValue(RootKey, SubKeyName, ValueName) {
  var objWMIService = WbemLocator.ConnectServer(strComputer, 'root\\default');
  var objReg = objWMIService.Get('StdRegProv');

  var objMethod = objReg.Methods_.Item('GetStringValue');
  var objInParam = objMethod.InParameters.SpawnInstance_();

  objInParam.hDefKey = RootKey;
  objInParam.sSubKeyName = SubKeyName;
  objInParam.sValueName = ValueName;

  var objOutParam = objReg.ExecMethod_('GetStringValue', objInParam);

  if (objOutParam.sValue != null) {
    return objOutParam.sValue;
  } else {
    return false;
  }
}

/**
 * 读取注册表值: array
 */
function getRegBinaryValue(RootKey, SubKeyName, ValueName) {
  var objWMIService = WbemLocator.ConnectServer(strComputer, 'root\\default');
  var objReg = objWMIService.Get('StdRegProv');

  var objMethod = objReg.Methods_.Item('GetBinaryValue');
  var objInParam = objMethod.InParameters.SpawnInstance_();

  objInParam.hDefKey = RootKey;
  objInParam.sSubKeyName = SubKeyName;
  objInParam.sValueName = ValueName;

  var objOutParam = objReg.ExecMethod_('GetBinaryValue', objInParam);

  if (objOutParam.uValue != null) {
    return new VBArray( objOutParam.uValue ).toArray();
  } else {
    return false;
  }
}


/**
 * 设置 x86, x64 正确的注册表节点
 */
function setRightRegNodePath(strKeyPath, OSArchitecture) {
  var nodeName = {
    "x86": "",
    "x64": "WOW6432Node"
  };
  var strNode = nodeName[OSArchitecture]
                ? '\\' + nodeName[OSArchitecture]
                : nodeName[OSArchitecture];

  var re = /WOW\d{4,}Node/i;

  return (
    strKeyPath.replace(/([A-Za-z]+)\\(\w+)(.*)/, function($0, $1, $2, $3) {
      return $1 +
        (
          $2.replace(re, '')
            ? strNode + '\\' + $2
            : strNode
        ) +
        $3;
    })
  );
}


/**
 * 提示手动退出一个程序
 */
function exitProgram(appName) {
  var objWMIService = WbemLocator.ConnectServer(strComputer, 'root\\cimv2');

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

    if (!isEmpty(app)) {
      vbMsgBox('请退出 ' + appName + ' 后继续 ...', 48, '提示信息');
    }

  } while (!isEmpty(app));
}


/**
 * 获取运行中程序的路径
 */
function getProcessAppProperties(appName) {
  var objWMIService = WbemLocator.ConnectServer(strComputer, 'root\\cimv2');
  var colItems = objWMIService.ExecQuery("Select * from Win32_Process where Name='" + appName + "'");

  var app = new Object();
  var enumItems = new Enumerator(colItems);

  for (; !enumItems.atEnd(); enumItems.moveNext()) {
    var properties = new Enumerator(enumItems.item().Properties_);

    for (; !properties.atEnd(); properties.moveNext()) {
      var name = properties.item().name;
      var value = properties.item().value;
      app[name] = value ? value : null;
    }
  }

  return isEmpty(app) ? false : formatPath(app.ExecutablePath);
}


/**
 * 路径状态信息
 */
function checkWindowsPathTips(str) {
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
function checkWindowsPath(str) {
  var re_h = /((.*)(?:\:)\\?)(.*)/;

  // Windows 默认文件名不能包括的字符 \/:*?"<>|
  var re_b = /[^`~!@#\$%\^&\(\)-\+=\[\]\{\};',\.\\\w\s]/g;

  var re_s = /[A-Za-z]/;

  return (
    re_h.test(str)
      ? str.replace(re_h, function($0, $1, $2, $3) {
        return (
          $2.length == 1 && re_s.test($2)
            ? $3.match(re_b)
              ? '{"status": 200, "string": "' + $3.match(re_b) + '"}'
              : '{"status": 300, "string": "' + $1 + $3 + '"}'
            : '{"status": 100, "string": "' + $1 + '"}'
        );
      })
      : '{"status": 0, "string": "' + str + '"}'
  );
}


/**
 * 格式化路径
 */
function formatPath(path, backslash) {
  path = !!path ? path : '';

  var len = path.length;
  var lastChar = len > 0 ? path.substring((len - 1), len) : '';

  if (lastChar == '\\') {
    path = path.substring(0, path.lastIndexOf('\\'));
  }

  return backslash ? path + '\\' : path;
}


/**
 * 获取时间
 */
function getDateTime() {
  var D = new Date();
  var y, m, d, H, i, s;

  y = D.getFullYear();
  m = D.getMonth() + 1;
  d = D.getDate();
  H = D.getHours();
  i = D.getMinutes();
  s = D.getSeconds();

  m = m < 10 ? '0' + m : m;
  d = d < 10 ? '0' + d : d;
  H = H < 10 ? '0' + H : H;
  i = i < 10 ? '0' + i : i;
  s = s < 10 ? '0' + s : s;

  return (
    y + "-" +
    m + "-" +
    d + ' ' +
    H + ":" +
    i + ":" +
    s
  );
}
