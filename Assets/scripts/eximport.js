var eximport = {
  'export': function(o) {
    msg = '即将开始备份 Steam 用户数据，请稍候...';
    log(msg);

    exitProgram('Steam.exe');
    exitProgram('dota2lauch.exe');

    o.stop().animate({width:'78%'},900);

    DeleteFolder(UserProfiles);
    CreateFolder(UserProfiles);

    fso.CopyFolder(SteamInstallPath + '\\config', UserProfiles + '\\');
    fso.CopyFolder(SteamInstallPath + '\\userdata', UserProfiles + '\\');
    fso.CopyFile(SteamInstallPath + '\\ssfn*', UserProfiles + '\\');

    msg = '正在打包 Steam 用户数据，请稍候...';
    log(msg);

    var outputFileName  = 'Steam_UserProfiles';
    var outputFile      = parentFolder + '\\' + outputFileName + '.exe';
    var makefile        = parentFolder + '\\make.vbs';
    var parameters      = '"' + outputFileName + '" "SILENT"';

    WshShell.Run('wscript.exe "' + makefile + '" ' + parameters, 1, true);

    // https://msdn.microsoft.com/en-us/library/windows/desktop/bb787866(v=vs.85).aspx
    // https://msdn.microsoft.com/en-us/library/windows/desktop/ms723207(v=vs.85).aspx
    var objShell = new ActiveXObject('Shell.Application');
    var objFolder = new Object;

    objFolder = objShell.NameSpace(WshShell.SpecialFolders('Desktop'));

    if (objFolder != null) {
      objFolder.CopyHere(outputFile, 8);
      DeleteFile(outputFile);
    }

    msg = 'Steam 用户数据备份完成，请在桌面查看 :）';
    log(msg);

    o.stop().animate({width:'100%'},450);
  },
  'import': function(o) {
    msg = '即将开始配置 Steam 用户数据，请稍候...';
    log(msg);

    if (fso.FolderExists(UserProfiles)) {
      var f = fso.GetFolder(UserProfiles);
      var size = (f.Size/1024/1024).toFixed(2); // MB

      if (size > 0) {
        exitProgram('Steam.exe');
        exitProgram('dota2lauch.exe');

        fso.CopyFolder(UserProfiles + '\\config', SteamInstallPath + '\\');
        fso.CopyFolder(UserProfiles + '\\userdata', SteamInstallPath + '\\');
        fso.CopyFile(UserProfiles + '\\ssfn*', SteamInstallPath + '\\');

        setRegStringValue(HKEY_CURRENT_USER, SubKeyPath, 'AutoLoginUser', this._lastLoginUser());

        msg = 'Steam 用户数据配置完成，即将启动 Steam 客户端 :）';
        log(msg);

        o.stop().animate({width:'100%'},500);

        if (fso.FileExists(SteamInstallPath + '\\Steam.exe')) {
          WshShell.Run('"' + SteamInstallPath + '\\Steam.exe"', 1, false);
        } else if (fso.FileExists(SteamInstallPath + '\\dota2lauch.exe')) {
          WshShell.Run('"' + SteamInstallPath + '\\dota2lauch.exe"', 1, false);
        }
      } else {
        msg = 'Steam 用户数据无效。';
        log(msg, 1);
      }
    } else {
      msg = '还没有 Steam 用户数据。';
      log(msg, 1);
    }
  },
  '_lastLoginUser': function() {
    if (!String.prototype.trim) {
      String.prototype.trim = function() {
        return this.replace(/(^\s+)|(\s+$)/g, '');
      };
    }

    var user = {
      AccountName: '',
      Timestamp: 0
    };

    var text = readVDFContent(UserProfiles + '\\config\\loginusers.vdf');
    var users = VDF.parse(text).users;

    for (var uid in users) {
      var AccountName = users[uid]['AccountName'];
      var Timestamp = parseInt(users[uid]['Timestamp']);

      if (user.Timestamp < Timestamp) {
        user.AccountName = AccountName;
        user.Timestamp = Timestamp;
      }
    }

    return user.AccountName;
  }
};