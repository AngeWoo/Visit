function doGet(e) {
    try {
        var action = e && e.parameter && e.parameter.action ? e.parameter.action : '';
        
        if (action === 'get_config') {
            return jsonOut(getMailConfig());
        }
        
        if (action === 'save_config') {
            var recipients = e.parameter.recipients || '';
            var time = e.parameter.time || '08:00';
            var result = saveMailConfig(recipients, time);
            return jsonOut(result);
        }
        
        if (action === 'test_mail') {
            var recipients = e.parameter.recipients || '';
            var result = triggerTestEmail(recipients);
            return jsonOut(result);
        }
        
        if (action === 'send_date_mail') {
            var recipients = e.parameter.recipients || '';
            var targetDate = e.parameter.date || '';
            var result = triggerTestEmail(recipients, targetDate);
            return jsonOut(result);
        }

        // Default: fetch CSV sheet data
        var defaultGid = '1401484943';
        var targetGid = e && e.parameter && e.parameter.gid ? e.parameter.gid : defaultGid;
        var url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSG6BO-kJ_GZoqn-JbMLhC_mDmlJ-q_5eWL4gUFnpoIrfvFf0iJ2uk4r0eQGZ9sfFVqL5Dx_UrEVOjI/pub?output=csv&gid=' + targetGid;
        var response = UrlFetchApp.fetch(url);
        var csvText = response.getContentText();
        var dataRows = Utilities.parseCsv(csvText);
        return jsonOut({ status: 'success', data: dataRows });
        
    } catch(err) {
        return jsonOut({ status: 'error', message: err.toString() });
    }
}

function jsonOut(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
    // Fallback for POST if ever called
    return doGet(e);
}

/* -------------------------------------
 *  Email Scheduling Automation Module 
 * ------------------------------------- */

function getMailConfig() {
    var props = PropertiesService.getScriptProperties();
    var hasTrigger = false;
    var triggerCount = 0;
    try {
        var triggers = ScriptApp.getProjectTriggers();
        for (var i = 0; i < triggers.length; i++) {
            var handler = triggers[i].getHandlerFunction();
            if (handler === 'dailySendMailTask' || handler === 'hourlyCheckAndSend') {
                hasTrigger = true;
                triggerCount++;
            }
        }
    } catch(e) {}
    return {
        recipients: props.getProperty('MAIL_RECIPIENTS') || 'ange.wu@ycgroup.tw',
        time: props.getProperty('MAIL_TIME') || '16:50',
        triggerActive: hasTrigger,
        triggerCount: triggerCount,
        lastSent: props.getProperty('LAST_SENT_DATETIME') || props.getProperty('LAST_SENT_DATE') || '尚未寄出'
    };
}

function saveMailConfig(recipients, timeString) {
    try {
        var props = PropertiesService.getScriptProperties();
        if (recipients) props.setProperty('MAIL_RECIPIENTS', recipients);
        if (timeString) props.setProperty('MAIL_TIME', timeString);

        // 自動建立每日精確時間觸發器
        var triggerMsg = ensureDailyTrigger(timeString);

        return { success: true, message: '設定已儲存！收件人: ' + recipients + '，每日 ' + timeString + ' 自動寄出。' + triggerMsg };
    } catch(e) {
        return { success: false, error: e.toString() };
    }
}

/**
 * 根據設定的時間建立每日精確觸發器。
 * 會先刪除舊的排程觸發器，再建立新的。
 */
function ensureDailyTrigger(timeString) {
    try {
        // 刪除所有舊的排程觸發器
        var triggers = ScriptApp.getProjectTriggers();
        for (var i = 0; i < triggers.length; i++) {
            var handler = triggers[i].getHandlerFunction();
            if (handler === 'dailySendMailTask' || handler === 'hourlyCheckAndSend') {
                ScriptApp.deleteTrigger(triggers[i]);
            }
        }

        // 解析時間
        var parts = timeString.split(':');
        var hour = parseInt(parts[0], 10);
        var minute = parseInt(parts[1], 10) || 0;

        // 建立週一到週五的觸發器
        var weekdays = [
            ScriptApp.WeekDay.MONDAY,
            ScriptApp.WeekDay.TUESDAY,
            ScriptApp.WeekDay.WEDNESDAY,
            ScriptApp.WeekDay.THURSDAY,
            ScriptApp.WeekDay.FRIDAY
        ];

        for (var i = 0; i < weekdays.length; i++) {
            ScriptApp.newTrigger('dailySendMailTask')
                .timeBased()
                .onWeekDay(weekdays[i])
                .atHour(hour)
                .nearMinute(minute)
                .create();
        }

        return '（已建立週一至週五 ' + timeString + ' 排程觸發器，分鐘精度 ±15 分鐘）';
    } catch(e) {
        return '（觸發器建立失敗: ' + e.toString() + '，請在 GAS 編輯器手動執行 setupDailyTrigger）';
    }
}

/**
 * [手動執行] 建立每日排程觸發器。
 * 會讀取已儲存的 MAIL_TIME 設定，建立精確時間觸發器。
 * 也可從網頁「儲存設定」自動建立，不需手動執行。
 */
function setupDailyTrigger() {
    var conf = getMailConfig();
    var result = ensureDailyTrigger(conf.time);
    Logger.log(result);
}

// 保留舊函式名稱相容
function setupHourlyTrigger() {
    setupDailyTrigger();
}

/**
 * 每小時自動執行，檢查現在是否為設定的寄信時間。
 * 若當前小時 = 設定小時，且今天尚未寄過，就寄出報表。
 */
function hourlyCheckAndSend() {
    var conf = getMailConfig();
    if (!conf.recipients || !conf.time) return;

    var parts = conf.time.split(':');
    var targetHour = parseInt(parts[0], 10);
    if (isNaN(targetHour)) return;

    var now = new Date();
    var tz = 'Asia/Taipei';
    var currentHour = parseInt(Utilities.formatDate(now, tz, 'H'), 10);
    var todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

    // 檢查現在是否為設定的小時
    if (currentHour !== targetHour) return;

    // 檢查今天是否已寄過（避免重複寄信）
    var props = PropertiesService.getScriptProperties();
    var lastSent = props.getProperty('LAST_SENT_DATE') || '';
    if (lastSent === todayStr) return;

    // 寄出報表
    try {
        var htmlBody = fetchStatsHtml();
        MailApp.sendEmail({
            to: conf.recipients,
            name: '炎洲訪客管理',
            subject: '炎洲集團各廠區 - ' + todayStr + ' 訪客明細',
            htmlBody: htmlBody
        });
        props.setProperty('LAST_SENT_DATE', todayStr);
        Logger.log('已於 ' + todayStr + ' ' + currentHour + ':00 寄出報表');
    } catch(e) {
        Logger.log('寄信失敗: ' + e.toString());
    }
}

function fetchStatsHtml(targetDateStr) {
    var sheetsInfo = [
        { name: '內湖', gid: '1401484943' },
        { name: '楊梅', gid: '930740199' },
        { name: '彰濱薄膜', gid: '412432769' },
        { name: '彰濱膠帶', gid: '545698913' }
    ];

    var now = new Date();
    var tz = 'Asia/Taipei';
    
    var filterY, filterM, filterD;
    if (targetDateStr) {
        var dp = targetDateStr.split('-');
        filterY = parseInt(dp[0], 10);
        filterM = parseInt(dp[1], 10);
        filterD = parseInt(dp[2], 10);
    } else {
        filterY = now.getFullYear();
        filterM = now.getMonth() + 1;
        filterD = now.getDate();
    }

    var dateLabel = filterY + '年' + filterM + '月' + filterD + '日';

    var headerStyle = 'background:#004e92;color:white;padding:8px 10px;text-align:left;font-size:12px;white-space:nowrap;';
    var tdStyle = 'padding:7px 8px;border:1px solid #e2e8f0;font-size:12px;white-space:nowrap;';
    var tblStyle = 'border-collapse:collapse;table-layout:fixed;width:1200px;font-family:sans-serif;margin-bottom:20px;';
    var colWidths = [40, 180, 80, 120, 130, 110, 110, 80, 170, 100];

    var message = '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="font-family:sans-serif;"><tr><td align="center"><table cellpadding="0" cellspacing="0" border="0" width="1200" style="font-family:sans-serif;max-width:1200px;"><tr><td>';
    message += '<h2 style="color:#004e92;border-bottom:3px solid #004e92;padding-bottom:8px;font-size:18px;">炎洲集團 - ' + dateLabel + ' 今日訪客詳細資料</h2>';

    for (var i = 0; i < sheetsInfo.length; i++) {
        var sheetObj = sheetsInfo[i];
        var todayRows = [];

        try {
            var url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSG6BO-kJ_GZoqn-JbMLhC_mDmlJ-q_5eWL4gUFnpoIrfvFf0iJ2uk4r0eQGZ9sfFVqL5Dx_UrEVOjI/pub?output=csv&gid=' + sheetObj.gid;
            var dataRows = Utilities.parseCsv(UrlFetchApp.fetch(url).getContentText());

            var [yLabel, mLabel, dLabel] = targetDateStr ? targetDateStr.split('-') : [filterY, filterM, filterD];
            var matchDatePrefix = yLabel + '/' + parseInt(mLabel) + '/' + parseInt(dLabel);

            var header = dataRows[0] || [];
            var idxName = -1, idxPhone = -1, idxVCompany = -1, idxTCompany = -1, idxTUnit = -1, idxTarget = -1, idxReason = -1, idxLeave = -1;
            for (var k=0; k<header.length; k++) {
              var h = header[k];
              if (h.indexOf('姓名') > -1 && h.indexOf('被訪') === -1) idxName = k;
              if (h.indexOf('手機') > -1 || h.indexOf('碼') > -1) idxPhone = k;
              if (h.indexOf('您的公司') > -1) idxVCompany = k;
              else if (h.indexOf('公司名稱') > -1) idxVCompany = k;
              else if ((h.indexOf('訪客') > -1 || h.indexOf('來訪') > -1) && h.indexOf('公司') > -1) idxVCompany = k;
              else if (h.indexOf('公司') > -1 && h.indexOf('欲拜訪') === -1 && h.indexOf('拜訪公司') === -1 && h.indexOf('單位') === -1) idxVCompany = k;
              if (h.indexOf('欲拜訪公司') > -1) idxTCompany = k;
              if (h.indexOf('欲拜訪單位') > -1) idxTUnit = k;
              if ((h.indexOf('被訪') > -1 && h.indexOf('人') > -1) || h.indexOf('對象') > -1 || h.indexOf('受訪者') > -1) idxTarget = k;
              if (h.indexOf('事由') > -1 || h.indexOf('原因') > -1) idxReason = k;
              if (h.indexOf('離') > -1 && h.indexOf('時間') > -1) idxLeave = k;
            }

            for (var j = 1; j < dataRows.length; j++) {
                var row = dataRows[j];
                var ts = row[0] || '';
                var match = ts.match(/(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
                
                if (match && 
                    parseInt(match[1], 10) === filterY && 
                    parseInt(match[2], 10) === filterM && 
                    parseInt(match[3], 10) === filterD) {
                    
                    // Normalize row for the table display (FULL COLUMNS)
                    var normRow = [
                      row[0], // 0: Timestamp (detailed)
                      (idxName > -1 ? row[idxName] : (row[3]||'-')), // 1: Name
                      (idxPhone > -1 ? row[idxPhone] : (row[4]||'-')), // 2: Phone
                      (idxVCompany > -1 ? row[idxVCompany] : (row[2]||'-')), // 3: Visitor Company
                      (idxTCompany > -1 ? row[idxTCompany] : (row[5]||'-')), // 4: Target Company
                      (idxTUnit > -1 ? row[idxTUnit] : (row[6]||'-')), // 5: Target Unit
                      (idxTarget > -1 ? row[idxTarget] : (row[8]||'-')), // 6: Target Person
                      (idxReason > -1 ? row[idxReason] : (row[7]||'-')), // 7: Reason
                      (idxLeave > -1 ? row[idxLeave] : (row[1]||'-')) // 8: Leave
                    ];
                    todayRows.push(normRow);
                }
            }
        } catch(e) {
            todayRows = [];
        }

        message += '<h3 style="background:#e8f0fe;padding:10px 14px;border-left:5px solid #004e92;margin:20px 0 6px;font-size:15px;">🏭 ' + sheetObj.name + '　今日訪客 ' + todayRows.length + ' 筆</h3>';

        if (todayRows.length === 0) {
            message += '<p style="color:#94a3b8;font-size:13px;padding:0 14px;">該日無訪客記錄</p>';
        } else {
            var colHeaders = ['#', '填表時間', '姓名', '手機', '訪客公司', '拜訪公司', '拜訪單位', '受訪者', '事由', '離廠時間'];
            message += '<table style="' + tblStyle + '" cellpadding="0" cellspacing="0" border="0">';
            message += '<colgroup>';
            for (var c = 0; c < colWidths.length; c++) {
                message += '<col width="' + colWidths[c] + '" style="width:' + colWidths[c] + 'px;">';
            }
            message += '</colgroup>';
            message += '<tr>';
            for (var c = 0; c < colHeaders.length; c++) {
                var align = (c === 0) ? 'text-align:center;' : '';
                message += '<th width="' + colWidths[c] + '" style="' + headerStyle + align + 'width:' + colWidths[c] + 'px;">' + colHeaders[c] + '</th>';
            }
            message += '</tr>';

            for (var r = 0; r < todayRows.length; r++) {
                var row = todayRows[r];
                var bg = (r % 2 === 0) ? '#f8fafc' : '#ffffff';
                var vals = [r + 1, row[0] || '-', row[1] || '-', row[2] || '-', row[3] || '-', row[4] || '-', row[5] || '-', row[6] || '-', row[7] || '-', row[8] || '-'];
                message += '<tr style="background:' + bg + '">';
                for (var c = 0; c < vals.length; c++) {
                    var extra = '';
                    if (c === 0) extra = 'text-align:center;color:#94a3b8;';
                    if (c === 2) extra = 'font-weight:bold;';
                    message += '<td width="' + colWidths[c] + '" style="' + tdStyle + extra + 'width:' + colWidths[c] + 'px;">' + vals[c] + '</td>';
                }
                message += '</tr>';
            }
            message += '</table>';
        }
    }

    message += '<p style="color:#94a3b8;font-size:11px;margin-top:20px;">本郵件由系統自動產生 · ' + Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm') + '</p>';
    message += '</td></tr></table></td></tr></table>';
    return message;
}

/**
 * 每日精確排程觸發 — 直接寄出報表（不需再比對時間）。
 */
function dailySendMailTask() {
    var conf = getMailConfig();
    if (!conf.recipients) return;

    var now = new Date();
    var tz = 'Asia/Taipei';
    var todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

    // 避免同一天重複寄信
    var props = PropertiesService.getScriptProperties();
    var lastSent = props.getProperty('LAST_SENT_DATE') || '';
    if (lastSent === todayStr) return;

    try {
        var htmlBody = fetchStatsHtml();
        MailApp.sendEmail({
            to: conf.recipients,
            name: '炎洲訪客管理',
            subject: '炎洲集團各廠區 - ' + todayStr + ' 訪客明細',
            htmlBody: htmlBody
        });
        var nowTimeStr = Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm');
        props.setProperty('LAST_SENT_DATE', todayStr);
        props.setProperty('LAST_SENT_DATETIME', nowTimeStr);
        Logger.log('已於 ' + nowTimeStr + ' 寄出報表');
    } catch(e) {
        Logger.log('寄信失敗: ' + e.toString());
    }
}

// 保留供舊觸發器相容
function scheduledSendMailTask() {
    dailySendMailTask();
}

function triggerTestEmail(recipientsStr, targetDateStr) {
    try {
        var conf = getMailConfig();
        var targetRecipients = recipientsStr || conf.recipients;
        var htmlBody = fetchStatsHtml(targetDateStr);
        
        var now = new Date();
        var tz = 'Asia/Taipei';
        var dateStr = targetDateStr || Utilities.formatDate(now, tz, 'yyyy-MM-dd');
        var subject = '【手動補發】炎洲集團各廠區 - ' + dateStr + ' 訪客明細';

        MailApp.sendEmail({
            to: targetRecipients,
            name: '炎洲訪客管理',
            subject: subject,
            htmlBody: htmlBody
        });

        var props = PropertiesService.getScriptProperties();
        var nowTimeStr = Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm');
        props.setProperty('LAST_SENT_DATETIME', nowTimeStr);

        return { success: true, message: '信件已成功發送！對象：' + targetRecipients };
    } catch(e) {
        return { success: false, error: e.toString() };
    }
}

function getSheetData(gid) {
    var defaultGid = '1401484943';
    var targetGid = gid || defaultGid;
    var url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSG6BO-kJ_GZoqn-JbMLhC_mDmlJ-q_5eWL4gUFnpoIrfvFf0iJ2uk4r0eQGZ9sfFVqL5Dx_UrEVOjI/pub?output=csv&gid=' + targetGid;

    var response = UrlFetchApp.fetch(url);
    var csvText = response.getContentText();
    var dataRows = Utilities.parseCsv(csvText);

    return {
        status: 'success',
        data: dataRows
    };
}
