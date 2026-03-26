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
    return {
        recipients: props.getProperty('MAIL_RECIPIENTS') || 'ange.wu@ycgroup.tw',
        time: props.getProperty('MAIL_TIME') || '08:00'
    };
}

function saveMailConfig(recipients, timeString) {
    try {
        var props = PropertiesService.getScriptProperties();
        if (recipients) props.setProperty('MAIL_RECIPIENTS', recipients);
        if (timeString) props.setProperty('MAIL_TIME', timeString);
        return { success: true, message: '設定已儲存！收件人: ' + recipients + '，時間: ' + timeString };
    } catch(e) {
        return { success: false, error: e.toString() };
    }
}

/**
 * [手動執行一次] 建立每日排程觸發器。
 * 請在 GAS 編輯器中選擇此函式然後執行一次即可，不需要重複執行。
 */
function createDailyTrigger() {
    // 刪除舊的同名觸發器
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'scheduledSendMailTask') {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }
    
    var conf = getMailConfig();
    var hour = 8;
    if (conf.time) {
        var parts = conf.time.split(':');
        if (parts.length > 0) hour = parseInt(parts[0], 10) || 8;
    }
    
    ScriptApp.newTrigger('scheduledSendMailTask')
        .timeBased()
        .everyDays(1)
        .atHour(hour)
        .nearMinute(0)
        .create();
        
    Logger.log('排程已建立，每天 ' + hour + ':00 寄信給 ' + conf.recipients);
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

    var headerStyle = 'background:#004e92;color:white;padding:8px 12px;text-align:left;font-size:12px;white-space:nowrap;';
    var tdStyle = 'padding:7px 10px;border:1px solid #e2e8f0;font-size:13px;white-space:nowrap;';
    var tblStyle = 'border-collapse:collapse;table-layout:auto;font-family:sans-serif;margin-bottom:20px;';

    var message = '<div style="font-family:sans-serif;max-width:100%;overflow-x:auto;margin:0 auto;">';
    message += '<h2 style="color:#004e92;border-bottom:3px solid #004e92;padding-bottom:8px;">炎洲集團 - ' + dateLabel + ' 今日訪客詳細資料</h2>';

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
              if (h.indexOf('公司名稱') > -1 || (h.indexOf('公司') > -1 && h.indexOf('欲拜訪') === -1)) idxVCompany = k;
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
            message += '<table style="' + tblStyle + '">';
            message += '<tr>';
            message += '<th style="' + headerStyle + 'width:36px;text-align:center;">#</th>';
            message += '<th style="' + headerStyle + '">填表時間</th>';
            message += '<th style="' + headerStyle + '">姓名</th>';
            message += '<th style="' + headerStyle + '">手機</th>';
            message += '<th style="' + headerStyle + '">訪客公司</th>';
            message += '<th style="' + headerStyle + '">拜訪公司</th>';
            message += '<th style="' + headerStyle + '">拜訪單位</th>';
            message += '<th style="' + headerStyle + '">受訪者</th>';
            message += '<th style="' + headerStyle + '">事由</th>';
            message += '<th style="' + headerStyle + '">離廠時間</th>';
            message += '</tr>';

            for (var r = 0; r < todayRows.length; r++) {
                var row = todayRows[r];
                var bg = (r % 2 === 0) ? '#f8fafc' : '#ffffff';
                message += '<tr style="background:' + bg + '">';
                message += '<td style="' + tdStyle + 'text-align:center;color: #94a3b8;">' + (r + 1) + '</td>';
                message += '<td style="' + tdStyle + ';white-space:nowrap;">' + (row[0] || '-') + '</td>';
                message += '<td style="' + tdStyle + ';font-weight:bold;white-space:nowrap;">' + (row[1] || '-') + '</td>';
                message += '<td style="' + tdStyle + ';white-space:nowrap;">' + (row[2] || '-') + '</td>';
                message += '<td style="' + tdStyle + ';white-space:nowrap;">' + (row[3] || '-') + '</td>';
                message += '<td style="' + tdStyle + ';white-space:nowrap;">' + (row[4] || '-') + '</td>';
                message += '<td style="' + tdStyle + ';white-space:nowrap;">' + (row[5] || '-') + '</td>';
                message += '<td style="' + tdStyle + ';white-space:nowrap;">' + (row[6] || '-') + '</td>';
                message += '<td style="' + tdStyle + '">' + (row[7] || '-') + '</td>';
                message += '<td style="' + tdStyle + ';white-space:nowrap;">' + (row[8] || '-') + '</td>';
                message += '</tr>';
            }
            message += '</table>';
        }
    }

    message += '<p style="color:#94a3b8;font-size:11px;margin-top:20px;">本郵件由系統自動產生 · ' + Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm') + '</p>';
    message += '</div>';
    return message;
}

function scheduledSendMailTask() {
    var conf = getMailConfig();
    if (!conf.recipients) return;
    var htmlBody = fetchStatsHtml();
    
    MailApp.sendEmail({
        to: conf.recipients,
        name: '炎洲防客管理系統',
        subject: '【系統報表】炎洲集團各廠區 - 最新訪客統計摘要',
        htmlBody: htmlBody
    });
}

function triggerTestEmail(recipientsStr, targetDateStr) {
    try {
        var conf = getMailConfig();
        var targetRecipients = recipientsStr || conf.recipients;
        var htmlBody = fetchStatsHtml(targetDateStr);
        
        var subject = targetDateStr 
            ? '【手動補發】炎洲集團各廠區 - ' + targetDateStr + ' 訪客明細'
            : '【手動測試】炎洲集團各廠區 - 報表發送演練';
            
        MailApp.sendEmail({
            to: targetRecipients,
            name: '炎洲防客管理',
            subject: subject,
            htmlBody: htmlBody
        });
        
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
