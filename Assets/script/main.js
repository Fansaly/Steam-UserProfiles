/**
 * 屏蔽鼠标右键
 */
jQuery_isTagName = function(e, whitelists) {
    e = $.event.fix(e);
    var target = e.target || e.srcElement;
    if (whitelists && $.inArray(target.tagName.toString().toLowerCase(), whitelists) == -1) {
        return false;
    }
    return true;
}
$(document).on('contextmenu', function(e){
    if (!jQuery_isTagName(e, ['input', 'textarea'])) {
        e.preventDefault();
        return false;
    }
    return true;
});


/**
 * 日志
 */
var $log = $('.log');
var log = function(o, error) {
    $log.animate({opacity: 0}, 90);
    setTimeout(function(){
        if (error) $progress.parent().animate({backgroundColor:progress['error']},160);

        $log.html(o).animate({opacity: 1}, 160);
    }, 120);
}


/**
 * Steam 相关变量
 */
var strKeyPath = 'Software\\Valve\\Steam';
var strValueName = 'InstallPath';
var SteamInstallPath = '';
var msg = '';

// 用户数据目录
var UserProfiles = parentFolder + '\\UserProfiles';

// 进度条
var $progress = $('.progress-bar'),
    progress = {
        'default': '#a0e9fd',
        // 'error': '#bf3737'
        'error': '#a94442'
    };


$('.button').on('click', function() {
    var $this = $(this),
        act = $this.attr('data-action');

    if ($this.data('status')) return false;

    $this.data('status', true);
    $this.addClass('active');

    stopAnimation();
    startAnimation();

    $progress.parent().animate({backgroundColor:progress['default']},90);
    $progress.stop().animate({width:'0%'},0);

    msg = '即将开始...';
    // vbPopup(msg, 2, '提示信息', 0 + 64 + 4096);
    log(msg);

    SteamInstallPath = getRegStringValue(HKEY_LOCAL_MACHINE, setRightRegNodePath(strKeyPath, OS.Architecture), strValueName);
    // SteamInstallPath = getProcessAppProperties('Steam.exe');

    if (SteamInstallPath) {

        if (fso.FolderExists(SteamInstallPath)) {

            switch (act) {
                case 'import':
                    _action['import']($progress);
                    break;
                case 'export':
                    _action['export']($progress);
                    break;
                default:
                    break;
            }

        } else {
            msg = 'Steam（国服）客户端的安装目录不存在  ╮（︶︿︶）╭';
            // vbMsgBox(msg + '\n"' + SteamInstallPath + '"', 16, '错误信息');
            log(msg, 1);
        }

    } else {
        msg = '没有找到 Steam（国服）客户端的安装记录 ╮（︶︿︶）╭';
        // vbMsgBox(msg, 64, '提示信息');
        log(msg, 1);
    }

    setTimeout(function(){
        $this.removeClass('active');
        $this.data('status', false);

        stopAnimation();
    }, 1000);
});


$('.copyleft a').on('click', function() {
    WshShell.Run('cmd.exe /C start ' + $(this).attr('href'), 0, false);
    return false;
});
