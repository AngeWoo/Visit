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

        if (action === 'force_send') {
            // 強制立即寄出（忽略今日已寄過的檢查）
            var recipients = e.parameter.recipients || '';
            var result = forceSendNow(recipients);
            return jsonOut(result);
        }

        if (action === 'clear_error') {
            var props = PropertiesService.getScriptProperties();
            props.deleteProperty('LAST_ERROR');
            props.deleteProperty('LAST_SENT_DATE');
            return jsonOut({ success: true, message: '已清除錯誤紀錄與今日寄送標記' });
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
            if (handler === 'scheduledCheckAndSend' || handler === 'dailySendMailTask' || handler === 'hourlyCheckAndSend') {
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
        lastSent: props.getProperty('LAST_SENT_DATETIME') || props.getProperty('LAST_SENT_DATE') || '尚未寄出',
        lastError: props.getProperty('LAST_ERROR') || '',
        lastTriggerRun: props.getProperty('LAST_TRIGGER_RUN') || ''
    };
}

function saveMailConfig(recipients, timeString) {
    try {
        var props = PropertiesService.getScriptProperties();
        if (recipients) props.setProperty('MAIL_RECIPIENTS', recipients);
        if (timeString) props.setProperty('MAIL_TIME', timeString);

        // 自動建立每10分鐘檢查觸發器
        var triggerMsg = ensureIntervalTrigger();

        return { success: true, message: '設定已儲存！收件人: ' + recipients + '，每日 ' + timeString + ' 自動寄出。' + triggerMsg };
    } catch(e) {
        return { success: false, error: e.toString() };
    }
}

/**
 * 建立每10分鐘執行一次的檢查觸發器（取代不穩定的 onWeekDay 觸發器）。
 * 只需要 1 個觸發器，每次執行檢查時間是否到了。
 */
function ensureIntervalTrigger() {
    try {
        // 刪除所有舊的排程觸發器
        var triggers = ScriptApp.getProjectTriggers();
        for (var i = 0; i < triggers.length; i++) {
            var handler = triggers[i].getHandlerFunction();
            if (handler === 'scheduledCheckAndSend' || handler === 'dailySendMailTask' || handler === 'hourlyCheckAndSend') {
                ScriptApp.deleteTrigger(triggers[i]);
            }
        }

        // 建立每10分鐘執行一次的觸發器
        ScriptApp.newTrigger('scheduledCheckAndSend')
            .timeBased()
            .everyMinutes(10)
            .create();

        return '（已建立每10分鐘自動檢查觸發器）';
    } catch(e) {
        var props = PropertiesService.getScriptProperties();
        props.setProperty('LAST_ERROR', '觸發器建立失敗: ' + e.toString());
        return '（觸發器建立失敗: ' + e.toString() + '，請在 GAS 編輯器手動執行 setupTrigger）';
    }
}

/**
 * [手動執行] 建立排程觸發器。
 * 在 GAS 編輯器選擇此函式並點選「執行」即可。
 */
function setupTrigger() {
    var result = ensureIntervalTrigger();
    Logger.log(result);
}

// 保留舊函式名稱相容
function setupDailyTrigger() { setupTrigger(); }
function setupHourlyTrigger() { setupTrigger(); }

/**
 * 每10分鐘自動執行：檢查是否為週一至週五的設定寄信時間。
 * 時間吻合且今日尚未寄過 → 寄出報表。
 */
function scheduledCheckAndSend() {
    var props = PropertiesService.getScriptProperties();
    var now = new Date();
    var tz = 'Asia/Taipei';
    var nowTimeStr = Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm');

    // 記錄觸發器確實有執行
    props.setProperty('LAST_TRIGGER_RUN', nowTimeStr);

    try {
        // 只在週一到週五執行
        var dayOfWeek = parseInt(Utilities.formatDate(now, tz, 'u'), 10); // 1=Mon ... 7=Sun
        if (dayOfWeek > 5) return; // 週六日不寄

        var conf = getMailConfig();
        if (!conf.recipients || !conf.time) return;

        // 比對現在時間是否在設定時間的 ±5 分鐘內
        var parts = conf.time.split(':');
        var targetHour = parseInt(parts[0], 10);
        var targetMinute = parseInt(parts[1], 10) || 0;
        var currentHour = parseInt(Utilities.formatDate(now, tz, 'H'), 10);
        var currentMinute = parseInt(Utilities.formatDate(now, tz, 'm'), 10);

        var targetTotal = targetHour * 60 + targetMinute;
        var currentTotal = currentHour * 60 + currentMinute;
        var diff = currentTotal - targetTotal;

        // 只在設定時間的 0~9 分鐘後寄出（配合每10分鐘觸發一次）
        if (diff < 0 || diff >= 10) return;

        var todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

        // 今天已寄過就跳過
        var lastSent = props.getProperty('LAST_SENT_DATE') || '';
        if (lastSent === todayStr) return;

        // 檢查配額
        var remaining = MailApp.getRemainingDailyQuota();
        if (remaining <= 0) {
            props.setProperty('LAST_ERROR', nowTimeStr + ' - MailApp 配額已用完');
            return;
        }

        // 寄出報表
        var htmlBody = fetchStatsHtml();
        MailApp.sendEmail({
            to: conf.recipients,
            name: '炎洲訪客管理',
            subject: '炎洲集團各廠區 - ' + todayStr + ' 訪客明細',
            htmlBody: htmlBody
        });
        props.setProperty('LAST_SENT_DATE', todayStr);
        props.setProperty('LAST_SENT_DATETIME', nowTimeStr);
        props.deleteProperty('LAST_ERROR');
        Logger.log('已於 ' + nowTimeStr + ' 寄出報表');
    } catch(e) {
        props.setProperty('LAST_ERROR', nowTimeStr + ' - ' + e.toString());
        Logger.log('排程寄信失敗: ' + e.toString());
    }
}

// 保留供舊觸發器相容
function hourlyCheckAndSend() { scheduledCheckAndSend(); }

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

// 保留供舊觸發器相容（如果還有舊的 dailySendMailTask 觸發器殘留）
function dailySendMailTask() {
    scheduledCheckAndSend();
}

// 保留供舊觸發器相容
function scheduledSendMailTask() {
    dailySendMailTask();
}

/**
 * 強制立即寄出報表（略過今日已寄過的檢查）
 */
function forceSendNow(recipientsStr) {
    try {
        var conf = getMailConfig();
        var targetRecipients = recipientsStr || conf.recipients;
        if (!targetRecipients) {
            return { success: false, error: '無收件人' };
        }

        var now = new Date();
        var tz = 'Asia/Taipei';
        var todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
        var nowTimeStr = Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm');

        var remaining = MailApp.getRemainingDailyQuota();
        if (remaining <= 0) {
            return { success: false, error: 'MailApp 每日配額已用完（剩餘:' + remaining + '）' };
        }

        var htmlBody = fetchStatsHtml();
        MailApp.sendEmail({
            to: targetRecipients,
            name: '炎洲訪客管理',
            subject: '炎洲集團各廠區 - ' + todayStr + ' 訪客明細',
            htmlBody: htmlBody
        });

        var props = PropertiesService.getScriptProperties();
        props.setProperty('LAST_SENT_DATE', todayStr);
        props.setProperty('LAST_SENT_DATETIME', nowTimeStr);
        props.deleteProperty('LAST_ERROR');

        return { success: true, message: '已強制寄出！對象：' + targetRecipients + '，配額剩餘：' + (remaining - 1) };
    } catch(e) {
        return { success: false, error: e.toString() };
    }
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
