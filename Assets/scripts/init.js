// vista 或以上版本系统以“管理员身份运行”
// if (OS.NT >= 60) {
//   var file = masterFile;
//   // 通过扩展名的大小写，来区分第几次运行 hta
//   if (file.substring(file.length-4) != '.HTA') {
//     // 以 runas 方式重新启动该 hta
//     Shell.ShellExecute('mshta.exe', '"' + file.substring(0, file.length-4) + '.HTA"', '', 'runas', 1);
//     // 退出当前 hta
//     window.close();
//     exit(0);
//   }
// }


/**
 * 获取任务栏注册表键值名: string
 */
function getTaskBarRegKeyValue(RootKey, SubKeyName) {
  var objWMIService = WbemLocator.ConnectServer(strComputer, 'root\\default');
  var objReg = objWMIService.Get('StdRegProv');
  var KeyName = null;

  var objMethod = objReg.Methods_.Item('EnumKey');
  var objInParam = objMethod.InParameters.SpawnInstance_();

  objInParam.hDefKey = RootKey;
  objInParam.sSubKeyName = SubKeyName;

  var objOutParam = objReg.ExecMethod_('EnumKey', objInParam);
  var arrSubKeys = new VBArray(objOutParam.sNames).toArray();

  for (var i = arrSubKeys.length - 1; i >= 0; i--) {
    if (!/[^StuckRects\d?]/i.test(arrSubKeys[i])) {
      KeyName = arrSubKeys[i];
      break;
    }
  }

  return KeyName ? KeyName : '';
}

// 获取任务栏的位置
function getTaskbarState() {
  var SubKeyMainPath = 'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer';
  var ValueName = 'Settings';
  var SubKeyPath, result;
  var state = new Array();

  SubKeyPath = SubKeyMainPath + '\\' + getTaskBarRegKeyValue(HKEY_CURRENT_USER, SubKeyMainPath);
  result = getRegBinaryValue(HKEY_CURRENT_USER, SubKeyPath, ValueName);

  for (var i = 0; i < result.length / 4; i++) {
    var begin = 4 * i;
    var end = begin + 4;
    state.push(result.slice(begin, end));
  }

  var taskbarAutoHide = !!(state[2][0] == 3);         // 任务栏自动隐藏状态
  var taskbarLocation = state[3][0];                  // 0-左 1-上 2-右 3-下

  var taskbarStartW = state[4][0] + state[4][1] * 256 // 开始按钮宽（也许）
  var taskbarStartH = state[5][0] + state[5][1] * 256 // 开始按钮高（也许）

  var taskbarLeft   = state[6][0] + state[6][1] * 256 // 任务栏左边界
  var taskbarTop    = state[7][0] + state[7][1] * 256 // 任务栏上边界
  var taskbarRight  = state[8][0] + state[8][1] * 256 // 任务栏右边界
  var taskbarBottom = state[9][0] + state[9][1] * 256 // 任务栏下边界

  // alert(
  //   '任务栏自动隐藏：' + taskbarAutoHide + '\n' +
  //   '任务栏所在位置：' + taskbarLocation + '\n\n' +
  //   '开始按钮宽（也许）：' + taskbarStartW + '\n' +
  //   '开始按钮高（也许）：' + taskbarStartH + '\n\n' +
  //   '任务栏左边界：' + taskbarLeft + '\n' +
  //   '任务栏上边界：' + taskbarTop + '\n' +
  //   '任务栏右边界：' + taskbarRight + '\n' +
  //   '任务栏下边界：' + taskbarBottom
  // );

  return [taskbarAutoHide, taskbarLocation];
}

// 应用程序窗口位置
function dialogLocation(w, h) {
  var screenWidth = window.screen.width;
  var screenHeight = window.screen.height;
  var workspaceWidth = window.screen.availWidth;
  var workspaceHeight = window.screen.availHeight;

  var left = (workspaceWidth - w) / 2;
  var top = (workspaceHeight - h) / 2;

  var taskbarState = getTaskbarState();

  if (!taskbarState[0]) {
    if (taskbarState[1] == 0) {
      left += screenWidth - workspaceWidth;
    } else if (taskbarState[1] == 1) {
      top += screenHeight - workspaceHeight;
    }
  }

  while (true) {
    try {
      window.resizeTo(w, h);
      window.moveTo(left, top);
      break;
    } catch (e) {
      continue;
    }
  }

  return true;
}

dialogLocation(600, 328);
