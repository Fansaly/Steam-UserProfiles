// 当前文件完整路径
var masterFile   = location.pathname;
// 当前项目路径
var parentFolder = masterFile.substring(0, masterFile.lastIndexOf('\\'));
// alert(parentFolder + '\n' + masterFile);

var Assets       = parentFolder + '\\Assets';
var Config       = parentFolder + '\\Config';
var temp         = parentFolder + '\\temp';
var Tools        = parentFolder + '\\Tools';
var UserProfiles = parentFolder + '\\UserProfiles';
var make         = parentFolder + '\\make.vbs';

// 注册表键值
var HKEY_CLASSES_ROOT   = 0x80000000;
var HKEY_CURRENT_USER   = 0x80000001;
var HKEY_LOCAL_MACHINE  = 0x80000002;
var HKEY_USERS          = 0x80000003;

var strComputer         = '.';          // 本地计算机
var OS                  = new Object(); // 操作系统信息

// 文件读写标志
var ForReading          =  1; // 只读模式
var ForWriting          =  2; // 只写模式
var ForAppending        =  8; // 文件末尾追加
var TristateFalse       =  0; // ASCII 格式
var TristateTrue        = -1; // Unicode 格式
var TristateUseDefault  = -2; // 系统默认格式

// ActiveXObject
var fso         = new ActiveXObject('Scripting.FileSystemObject');
var Shell       = new ActiveXObject('Shell.Application');
var WshShell    = new ActiveXObject('WScript.Shell');
var WbemLocator = new ActiveXObject('WbemScripting.SWbemLocator');


OS = getOperatingSystem();
