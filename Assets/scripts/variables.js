var masterFile = decodeURI(location.pathname)
                  .replace(/\//g, '\\')
                  .replace(/(^\\+)/, '');

var masterDir = masterFile.substring(0, masterFile.lastIndexOf('\\'));

var assetsDir = masterDir + '\\Assets';
var configDir = masterDir + '\\Config';
var toolsDir  = masterDir + '\\Tools';
var tempDir   = masterDir + '\\temp';
var makefile  = masterDir + '\\make.vbs';

// 注册表键值
var HKEY_CLASSES_ROOT  = 0x80000000;
var HKEY_CURRENT_USER  = 0x80000001;
var HKEY_LOCAL_MACHINE = 0x80000002;
var HKEY_USERS         = 0x80000003;

// 文件读写标志
var ForReading         =  1; // 只读模式
var ForWriting         =  2; // 只写模式
var ForAppending       =  8; // 文件末尾追加
var TristateFalse      =  0; // ASCII 格式
var TristateTrue       = -1; // Unicode 格式
var TristateUseDefault = -2; // 系统默认格式

// ActiveXObject
var fso         = new ActiveXObject('Scripting.FileSystemObject');
var WshShell    = new ActiveXObject('WScript.Shell');
var AppShell    = new ActiveXObject('Shell.Application');
var WbemLocator = new ActiveXObject('WbemScripting.SWbemLocator');

// 本地计算机
var strComputer = '.';
// 操作系统信息
var OS          = getOperatingSystemInfo();
var uiScale     = window.screen.deviceXDPI / 96;

createFolder(tempDir);
