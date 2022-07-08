/**
 * 屏蔽右键菜单
 */
// function jQuery_isTagName(e, whitelists) {
//   e = $.event.fix(e);
//   var target = e.target || e.srcElement;
//   if (whitelists && $.inArray(target.tagName.toString().toLowerCase(), whitelists) == -1) {
//       return false;
//   }
//   return true;
// }
// $(document).on('contextmenu', function(e){
//   if (!jQuery_isTagName(e, ['input', 'textarea'])) {
//     e.preventDefault();
//     return false;
//   }
//   return true;
// });


var defaultFontSize = 24;
var fontSize = uiScale * defaultFontSize;
$('html').css({ 'font-size': fontSize });


// 进度条及颜色
var $progress = $('.progress-bar');
var progressColor = {
  'success': '#35a260', // 389e60
  'warning': '#8a6d3b',
  'normal': '#a0e9fd',
  'error': '#a94442' // bf3737
};

// 状态
var $log = $('.log');
var log = function(msg, status) {
  if (status == 'error') {
    $progress
      .parent()
      .animate({
        backgroundColor: progressColor.error
      }, 160);
  } else if (status == 'fail') {
    $progress
      .animate({
        backgroundColor: progressColor.error
      }, 200)
      .animate({width: '100%'}, 1e3);
  } else if (status == 'warning') {
    $progress
      .animate({
        backgroundColor: progressColor.warning
      }, 200)
      .animate({width: '100%'}, 1e3);
  } else if (status == 'success') {
    $progress
      .animate({width: '100%'}, 1e3);
  }

  $log.animate({opacity: 0}, 90);
  setTimeout(function() {
    $log.html(msg).animate({opacity: 1}, 160);
  }, 120);
};


// 数据 导入、导出
var eximport = {
  ex: function() {
    log('即将开始备份 Steam 用户数据，请稍候 …');

    $progress.stop().animate({width: '10%'}, 100);

    exitProgram('Steam.exe');
    exitProgram('dota2lauch.exe');

    log('正在打包 Steam 用户数据，请稍候 …');
    $progress.stop().animate({width: '78%'}, 500);

    var steamInstallPath = this.options.installPath;
    var userProfiles = this.options.userProfiles;

    deleteFolder(userProfiles);
    createFolder(userProfiles);

    fso.CopyFolder(steamInstallPath + '\\config', userProfiles + '\\');
    fso.CopyFolder(steamInstallPath + '\\userdata', userProfiles + '\\');
    fso.CopyFile(steamInstallPath + '\\ssfn*', userProfiles + '\\');

    var outputFileName = 'Steam_UserProfiles';
    var outputFile     = masterDir + '\\' + outputFileName + '.exe';
    var parameters     = '"' + outputFileName + '" "SILENT"';

    WshShell.Run('wscript.exe "' + makefile + '" ' + parameters, 1, true);

    var result = copyToDesktop(outputFile);

    if (result) {
      deleteFile(outputFile);
      log('Steam 用户数据备份完成，请在桌面查看 :）', 'success');
    } else {
      log('Steam 用户数据文件，复制到桌面失败 :(', 'fail');
    }
  },
  im: function() {
    log('即将开始配置 Steam 用户数据，请稍候 …');

    var steamInstallPath = this.options.installPath;
    var userProfiles = this.options.userProfiles;

    if (fso.FolderExists(userProfiles)) {
      var f = fso.GetFolder(userProfiles);
      var size = (f.Size/1024/1024).toFixed(2);

      if (size > 0) {
        exitProgram('Steam.exe');
        exitProgram('dota2lauch.exe');

        fso.CopyFolder(userProfiles + '\\config', steamInstallPath + '\\');
        fso.CopyFolder(userProfiles + '\\userdata', steamInstallPath + '\\');
        fso.CopyFile(userProfiles + '\\ssfn*', steamInstallPath + '\\');

        setRegStringValue(
          HKEY_CURRENT_USER,
          this.options.subKeyPath,
          'AutoLoginUser',
          this.lastLoginUser()
        );

        log('Steam 用户数据配置完成，即将启动 Steam 客户端 :）', 'success');

        if (fso.FileExists(steamInstallPath + '\\Steam.exe')) {
          WshShell.Run('"' + steamInstallPath + '\\Steam.exe"', 1, false);
        } else if (fso.FileExists(steamInstallPath + '\\dota2lauch.exe')) {
          WshShell.Run('"' + steamInstallPath + '\\dota2lauch.exe"', 1, false);
        }
      } else {
        log('Steam 用户数据无效。', 'error');
      }
    } else {
      log('还没有 Steam 用户数据。', 'error');
    }
  },
  lastLoginUser: function() {
    var loginusers = this.options.userProfiles + '\\config\\loginusers.vdf';
    var text;

    try {
      text = readUTF8Content(loginusers);
    } catch (e) {
      text = readNormalContent(loginusers);
    }

    var users = VDF.parse(text).users;

    var user = {
      AccountName: '',
      Timestamp: 0
    };

    for (var uid in users) {
      var AccountName = users[uid]['AccountName'];
      var Timestamp = parseInt(users[uid]['Timestamp']);

      if (user.Timestamp < Timestamp) {
        user.AccountName = AccountName;
        user.Timestamp = Timestamp;
      }
    }

    return user.AccountName;
  },
  init: function() {
    var userProfiles = masterDir + '\\UserProfiles';
    var subKeyPath = 'Software\\Valve\\Steam';
    var valueName = 'InstallPath';
    var installPath = '';

    // installPath = getProcessAppProperties('Steam.exe');
    installPath = getRegStringValue(
      HKEY_LOCAL_MACHINE,
      setRightRegNodePath(subKeyPath, OS.Architecture),
      valueName
    );

    if (!installPath) {
      log('没有找到 Steam 客户端的安装记录 ╮（︶︿︶）╭', 'error');
    } else if (!fso.FolderExists(installPath)) {
      log('Steam 客户端的安装目录不存在  ╮（︶︿︶）╭', 'error');
    } else {
      this.options = {
        subKeyPath: subKeyPath,
        installPath: installPath,
        userProfiles: userProfiles
      };

      log('准备就绪。');

      $('.button')
        .stop()
        .animate({opacity: 1}, 365)
        .removeClass('disabled');
    }
  }
};

// 初始化
setTimeout(function() {
  eximport.init();
}, 0);


// 导入、导出事件监听
$('.button').on('click', function() {
  var $this = $(this);
  var action = $this.attr('data-action');

  if ($this.parent().hasClass('executing') || $this.hasClass('disabled')) {
    return false;
  }

  $this.parent().addClass('executing');
  $this.stop().animate({opacity: .9}, 365).addClass('active');

  stopAnimation();
  startAnimation();

  $progress.parent().animate({backgroundColor: progressColor.normal}, 90);
  $progress.stop().animate({
    width: '0%',
    backgroundColor: progressColor.success
  }, 0);

  if (action === 'import') {
    eximport.im();
  } else if (action === 'export') {
    eximport.ex();
  }

  setTimeout(function() {
    stopAnimation();

    $this.stop().animate({opacity: 1}, 365).removeClass('active');
    $this.parent().removeClass('executing');
  }, 1e3);
});


$('.copyright a').on('click', function() {
  WshShell.Run('cmd.exe /C start ' + $(this).attr('href'), 0, false);
  return false;
});


// 清理
$(window).on('beforeunload', function() {
  var isSource = fso.FileExists(masterDir + '\\.git\\config');
  if (isSource) return;

  deleteFolder(eximport.options.userProfiles);

  deleteFolder(assetsDir);
  deleteFolder(configDir);
  deleteFolder(tempDir);
  deleteFolder(toolsDir);

  deleteFile(makefile);
  deleteFile(masterFile);
});
