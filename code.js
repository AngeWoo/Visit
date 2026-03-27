var TIME_ZONE = 'Asia/Taipei';
var DEFAULT_GID = '1401484943';
var PUBLIC_SHEET_BASE =
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vSG6BO-kJ_GZoqn-JbMLhC_mDmlJ-q_5eWL4gUFnpoIrfvFf0iJ2uk4r0eQGZ9sfFVqL5Dx_UrEVOjI/pub?output=csv&gid=';
var DEFAULT_RECIPIENTS = 'ange.wu@ycgroup.tw';
var DEFAULT_TIME = '16:50';
var MAIL_SENDER_NAME = '各廠區訪客系統';
var MAIL_SUBJECT_PREFIX = '各廠區訪客日報';
var SHEETS_INFO = [
    { name: '頭份', gid: '1401484943' },
    { name: '樹林', gid: '930740199' },
    { name: '觀音三廠', gid: '412432769' },
    { name: '觀音二廠', gid: '545698913' }
];

function doGet(e) {
    try {
        var action = getParam_(e, 'action', '');

        if (action === 'get_config') {
            return jsonOut(getMailConfig());
        }

        if (action === 'save_config') {
            return jsonOut(
                saveMailConfig(
                    getParam_(e, 'recipients', ''),
                    getParam_(e, 'time', DEFAULT_TIME)
                )
            );
        }

        if (action === 'test_mail') {
            return jsonOut(triggerTestEmail(getParam_(e, 'recipients', '')));
        }

        if (action === 'send_date_mail') {
            return jsonOut(
                triggerTestEmail(
                    getParam_(e, 'recipients', ''),
                    getParam_(e, 'date', '')
                )
            );
        }

        if (action === 'force_send') {
            return jsonOut(forceSendNow(getParam_(e, 'recipients', '')));
        }

        if (action === 'clear_error') {
            return jsonOut(clearMailStatus());
        }

        return jsonOut(getSheetData(getParam_(e, 'gid', DEFAULT_GID)));
    } catch (err) {
        return jsonOut({
            status: 'error',
            message: err && err.toString ? err.toString() : String(err)
        });
    }
}

function doPost(e) {
    return doGet(e);
}

function jsonOut(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
}

function getParam_(e, key, fallback) {
    if (e && e.parameter && e.parameter[key] !== undefined && e.parameter[key] !== null) {
        return e.parameter[key];
    }
    return fallback;
}

function getMailConfig() {
    var props = PropertiesService.getScriptProperties();
    var triggers = [];
    var triggerCount = 0;

    try {
        triggers = ScriptApp.getProjectTriggers();
    } catch (e) {
        triggers = [];
    }

    for (var i = 0; i < triggers.length; i++) {
        if (isMailTriggerHandler_(triggers[i].getHandlerFunction())) {
            triggerCount++;
        }
    }

    return {
        recipients: props.getProperty('MAIL_RECIPIENTS') || DEFAULT_RECIPIENTS,
        time: props.getProperty('MAIL_TIME') || DEFAULT_TIME,
        triggerActive: triggerCount > 0,
        triggerCount: triggerCount,
        lastSent: props.getProperty('LAST_SENT_DATETIME') || props.getProperty('LAST_SENT_DATE') || '尚未寄出',
        lastError: props.getProperty('LAST_ERROR') || '',
        lastTriggerRun: props.getProperty('LAST_TRIGGER_RUN') || '',
        lastManualSent: props.getProperty('LAST_MANUAL_SENT_DATETIME') || '',
        lastScheduledSent: props.getProperty('LAST_SCHEDULED_SENT_DATE') || ''
    };
}

function saveMailConfig(recipients, timeString) {
    try {
        var props = PropertiesService.getScriptProperties();
        var normalizedRecipients = normalizeRecipients_(recipients);
        var normalizedTime = normalizeTime_(timeString);
        var triggerMsg;

        props.setProperty('MAIL_RECIPIENTS', normalizedRecipients || DEFAULT_RECIPIENTS);
        props.setProperty('MAIL_TIME', normalizedTime);

        triggerMsg = ensureIntervalTrigger();

        return {
            success: true,
            message:
                '設定已儲存，收件人：' +
                (normalizedRecipients || DEFAULT_RECIPIENTS) +
                '，排程時間：' +
                normalizedTime +
                '。' +
                triggerMsg
        };
    } catch (e) {
        return { success: false, error: safeErrorMessage_(e) };
    }
}

function clearMailStatus() {
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty('LAST_ERROR');
    props.deleteProperty('LAST_SENT_DATE');
    props.deleteProperty('LAST_SENT_DATETIME');
    props.deleteProperty('LAST_TRIGGER_RUN');
    props.deleteProperty('LAST_MANUAL_SENT_DATE');
    props.deleteProperty('LAST_MANUAL_SENT_DATETIME');
    props.deleteProperty('LAST_SCHEDULED_SENT_DATE');
    props.deleteProperty('LAST_SCHEDULED_SENT_KEY');
    return { success: true, message: '寄信狀態與錯誤紀錄已清除。' };
}

function ensureIntervalTrigger() {
    try {
        var triggers = ScriptApp.getProjectTriggers();
        var i;

        for (i = 0; i < triggers.length; i++) {
            if (isMailTriggerHandler_(triggers[i].getHandlerFunction())) {
                ScriptApp.deleteTrigger(triggers[i]);
            }
        }

        ScriptApp.newTrigger('scheduledCheckAndSend')
            .timeBased()
            .everyMinutes(10)
            .create();

        return '已建立每 10 分鐘檢查一次的排程觸發器。';
    } catch (e) {
        PropertiesService.getScriptProperties().setProperty(
            'LAST_ERROR',
            '建立觸發器失敗：' + safeErrorMessage_(e)
        );
        return '建立觸發器失敗：' + safeErrorMessage_(e);
    }
}

function setupTrigger() {
    Logger.log(ensureIntervalTrigger());
}

function setupDailyTrigger() {
    setupTrigger();
}

function setupHourlyTrigger() {
    setupTrigger();
}

function scheduledCheckAndSend() {
    var props = PropertiesService.getScriptProperties();
    var now = new Date();
    var nowTimeStr = Utilities.formatDate(now, TIME_ZONE, 'yyyy/MM/dd HH:mm');

    props.setProperty('LAST_TRIGGER_RUN', nowTimeStr);

    try {
        var dayOfWeek = parseInt(Utilities.formatDate(now, TIME_ZONE, 'u'), 10);
        var conf;
        var parts;
        var targetHour;
        var targetMinute;
        var currentHour;
        var currentMinute;
        var targetTotal;
        var currentTotal;
        var diff;
        var todayStr;
        var scheduleSlotKey;
        var lastScheduledSlot;
        var remaining;
        var htmlBody;

        if (dayOfWeek > 5) {
            return;
        }

        conf = getMailConfig();
        if (!conf.recipients || !conf.time) {
            return;
        }

        parts = normalizeTime_(conf.time).split(':');
        targetHour = parseInt(parts[0], 10);
        targetMinute = parseInt(parts[1], 10);
        currentHour = parseInt(Utilities.formatDate(now, TIME_ZONE, 'H'), 10);
        currentMinute = parseInt(Utilities.formatDate(now, TIME_ZONE, 'm'), 10);

        targetTotal = targetHour * 60 + targetMinute;
        currentTotal = currentHour * 60 + currentMinute;
        diff = currentTotal - targetTotal;

        if (diff < 0 || diff >= 10) {
            return;
        }

        todayStr = Utilities.formatDate(now, TIME_ZONE, 'yyyy-MM-dd');
        scheduleSlotKey = todayStr + ' ' + conf.time;
        lastScheduledSlot = props.getProperty('LAST_SCHEDULED_SENT_KEY') || '';
        if (lastScheduledSlot === scheduleSlotKey) {
            return;
        }

        remaining = MailApp.getRemainingDailyQuota();
        if (remaining <= 0) {
            props.setProperty('LAST_ERROR', nowTimeStr + ' - MailApp 每日配額已用完');
            return;
        }

        htmlBody = fetchStatsHtml();
        MailApp.sendEmail({
            to: conf.recipients,
            name: MAIL_SENDER_NAME,
            subject: MAIL_SUBJECT_PREFIX + ' - ' + todayStr,
            htmlBody: htmlBody
        });

        props.setProperty('LAST_SCHEDULED_SENT_KEY', scheduleSlotKey);
        props.setProperty('LAST_SCHEDULED_SENT_DATE', todayStr);
        props.setProperty('LAST_SENT_DATE', todayStr);
        props.setProperty('LAST_SENT_DATETIME', nowTimeStr);
        props.deleteProperty('LAST_ERROR');
        Logger.log('排程郵件寄送完成：' + nowTimeStr);
    } catch (e) {
        props.setProperty('LAST_ERROR', nowTimeStr + ' - ' + safeErrorMessage_(e));
        Logger.log('排程寄信失敗：' + safeErrorMessage_(e));
    }
}

function hourlyCheckAndSend() {
    scheduledCheckAndSend();
}

function dailySendMailTask() {
    scheduledCheckAndSend();
}

function scheduledSendMailTask() {
    dailySendMailTask();
}

function forceSendNow(recipientsStr) {
    try {
        var conf = getMailConfig();
        var targetRecipients = normalizeRecipients_(recipientsStr || conf.recipients);
        var now = new Date();
        var todayStr = Utilities.formatDate(now, TIME_ZONE, 'yyyy-MM-dd');
        var nowTimeStr = Utilities.formatDate(now, TIME_ZONE, 'yyyy/MM/dd HH:mm');
        var remaining = MailApp.getRemainingDailyQuota();
        var htmlBody;
        var props;

        if (!targetRecipients) {
            return { success: false, error: '請先設定至少一位收件人。' };
        }

        if (remaining <= 0) {
            return { success: false, error: 'MailApp 每日配額已用完（剩餘：' + remaining + '）。' };
        }

        htmlBody = fetchStatsHtml();
        MailApp.sendEmail({
            to: targetRecipients,
            name: MAIL_SENDER_NAME,
            subject: MAIL_SUBJECT_PREFIX + ' - ' + todayStr,
            htmlBody: htmlBody
        });

        props = PropertiesService.getScriptProperties();
        props.setProperty('LAST_MANUAL_SENT_DATE', todayStr);
        props.setProperty('LAST_MANUAL_SENT_DATETIME', nowTimeStr);
        props.setProperty('LAST_SENT_DATETIME', nowTimeStr);
        props.deleteProperty('LAST_ERROR');

        return {
            success: true,
            message: '已強制寄出郵件給：' + targetRecipients + '。剩餘配額：約 ' + (remaining - 1)
        };
    } catch (e) {
        return { success: false, error: safeErrorMessage_(e) };
    }
}

function triggerTestEmail(recipientsStr, targetDateStr) {
    try {
        var conf = getMailConfig();
        var targetRecipients = normalizeRecipients_(recipientsStr || conf.recipients);
        var htmlBody = fetchStatsHtml(targetDateStr);
        var now = new Date();
        var dateStr = targetDateStr || Utilities.formatDate(now, TIME_ZONE, 'yyyy-MM-dd');
        var nowTimeStr = Utilities.formatDate(now, TIME_ZONE, 'yyyy/MM/dd HH:mm');

        if (!targetRecipients) {
            return { success: false, error: '請先設定至少一位收件人。' };
        }

        MailApp.sendEmail({
            to: targetRecipients,
            name: MAIL_SENDER_NAME,
            subject: '測試郵件 - ' + MAIL_SUBJECT_PREFIX + ' - ' + dateStr,
            htmlBody: htmlBody
        });

        PropertiesService.getScriptProperties().setProperty('LAST_SENT_DATETIME', nowTimeStr);

        return {
            success: true,
            message: '測試郵件已寄出給：' + targetRecipients
        };
    } catch (e) {
        return { success: false, error: safeErrorMessage_(e) };
    }
}

function fetchStatsHtml(targetDateStr) {
    var dateInfo = normalizeTargetDate_(targetDateStr);
    var titleDate = dateInfo.year + '/' + dateInfo.month + '/' + dateInfo.day;
    var nowText = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy/MM/dd HH:mm');
    var message = '';
    var i;

    message += '<div style="font-family:Arial,Microsoft JhengHei,sans-serif;color:#1f2937;line-height:1.6;">';
    message +=
        '<h2 style="margin:0 0 16px;color:#004e92;border-bottom:3px solid #004e92;padding-bottom:8px;">' +
        escapeHtml_(MAIL_SUBJECT_PREFIX) +
        ' - ' +
        escapeHtml_(titleDate) +
        '</h2>';

    for (i = 0; i < SHEETS_INFO.length; i++) {
        message += buildFactorySectionHtml_(SHEETS_INFO[i], dateInfo);
    }

    message +=
        '<p style="margin-top:24px;color:#64748b;font-size:12px;">系統產生時間：' +
        escapeHtml_(nowText) +
        '</p>';
    message += '</div>';

    return message;
}

function buildFactorySectionHtml_(sheetInfo, dateInfo) {
    var rows;
    var tableRows;
    var html = '';
    var i;
    var displayRows;

    try {
        rows = normalizeFactoryRows_(sheetInfo.gid, dateInfo);
    } catch (e) {
        rows = [];
    }

    html +=
        '<h3 style="background:#e8f0fe;padding:10px 14px;border-left:5px solid #004e92;margin:20px 0 8px;font-size:15px;">' +
        escapeHtml_(sheetInfo.name) +
        ' | 今日筆數：' +
        rows.length +
        '</h3>';

    if (!rows.length) {
        html += '<p style="margin:0 0 16px 14px;color:#94a3b8;font-size:13px;">本日無符合資料。</p>';
        return html;
    }

    tableRows = '';
    displayRows = ['#', '時間', '姓名', '手機', '訪客公司', '拜訪公司', '拜訪單位', '受訪者', '事由', '離場時間'];

    html += '<table style="border-collapse:collapse;width:100%;min-width:1100px;font-size:12px;margin-bottom:16px;">';
    html += '<thead><tr style="background:#004e92;color:#ffffff;">';
    for (i = 0; i < displayRows.length; i++) {
        tableRows += '';
        html +=
            '<th style="padding:8px;border:1px solid #dbeafe;text-align:left;white-space:nowrap;">' +
            escapeHtml_(displayRows[i]) +
            '</th>';
    }
    html += '</tr></thead><tbody>';

    for (i = 0; i < rows.length; i++) {
        html += '<tr style="background:' + (i % 2 === 0 ? '#f8fafc' : '#ffffff') + ';">';
        html += buildCell_(i + 1, true);
        html += buildCell_(rows[i].time);
        html += buildCell_(rows[i].name);
        html += buildCell_(rows[i].phone);
        html += buildCell_(rows[i].visitorCompany);
        html += buildCell_(rows[i].targetCompany);
        html += buildCell_(rows[i].targetUnit);
        html += buildCell_(rows[i].targetPerson);
        html += buildCell_(rows[i].reason);
        html += buildCell_(rows[i].leaveTime);
        html += '</tr>';
    }

    html += '</tbody></table>';
    return html;
}

function buildCell_(value, center) {
    var extraStyle = center ? 'text-align:center;' : '';
    return (
        '<td style="padding:8px;border:1px solid #e2e8f0;white-space:nowrap;' +
        extraStyle +
        '">' +
        escapeHtml_(value || '-') +
        '</td>'
    );
}

function normalizeFactoryRows_(gid, dateInfo) {
    var dataRows = fetchCsvRows_(gid);
    var header = dataRows[0] || [];
    var normalizedRows = [];
    var idx = detectColumnIndexes_(header);
    var i;

    for (i = 1; i < dataRows.length; i++) {
        var row = dataRows[i];
        var timestamp = safeCell_(row, 0);
        var rowDate = parseRowDate_(timestamp);

        if (!rowDate) {
            continue;
        }

        if (
            rowDate.year !== dateInfo.year ||
            rowDate.month !== dateInfo.month ||
            rowDate.day !== dateInfo.day
        ) {
            continue;
        }

        normalizedRows.push({
            time: timestamp,
            name: getValueByIndex_(row, idx.name, 3),
            phone: getValueByIndex_(row, idx.phone, 4),
            visitorCompany: getValueByIndex_(row, idx.visitorCompany, 2),
            targetCompany: getValueByIndex_(row, idx.targetCompany, 5),
            targetUnit: getValueByIndex_(row, idx.targetUnit, 6),
            targetPerson: getValueByIndex_(row, idx.targetPerson, 8),
            reason: getValueByIndex_(row, idx.reason, 7),
            leaveTime: getValueByIndex_(row, idx.leaveTime, 1)
        });
    }

    return normalizedRows;
}

function detectColumnIndexes_(header) {
    return {
        name: findHeaderIndex_(header, [
            ['姓名'],
            ['訪客', '姓名'],
            ['來訪', '姓名']
        ]),
        phone: findHeaderIndex_(header, [['手機'], ['電話']]),
        visitorCompany: findHeaderIndex_(header, [
            ['您的公司'],
            ['公司名稱'],
            ['訪客', '公司'],
            ['來訪', '公司']
        ]),
        targetCompany: findHeaderIndex_(header, [['欲拜訪公司'], ['拜訪公司']]),
        targetUnit: findHeaderIndex_(header, [['欲拜訪單位'], ['拜訪單位'], ['部門']]),
        targetPerson: findHeaderIndex_(header, [
            ['拜訪對象'],
            ['被訪者'],
            ['受訪者'],
            ['接待人']
        ]),
        reason: findHeaderIndex_(header, [['事由'], ['原因'], ['目的']]),
        leaveTime: findHeaderIndex_(header, [['離場'], ['離廠'], ['離開時間']])
    };
}

function findHeaderIndex_(header, candidateGroups) {
    var i;
    var j;
    var cell;

    for (i = 0; i < header.length; i++) {
        cell = String(header[i] || '');
        for (j = 0; j < candidateGroups.length; j++) {
            if (containsAllKeywords_(cell, candidateGroups[j])) {
                return i;
            }
        }
    }

    return -1;
}

function containsAllKeywords_(text, keywords) {
    var i;
    for (i = 0; i < keywords.length; i++) {
        if (String(text).indexOf(keywords[i]) === -1) {
            return false;
        }
    }
    return true;
}

function getValueByIndex_(row, detectedIndex, fallbackIndex) {
    if (detectedIndex > -1) {
        return safeCell_(row, detectedIndex);
    }
    return safeCell_(row, fallbackIndex);
}

function parseRowDate_(text) {
    var match = String(text || '').match(/(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
    if (!match) {
        return null;
    }
    return {
        year: parseInt(match[1], 10),
        month: parseInt(match[2], 10),
        day: parseInt(match[3], 10)
    };
}

function normalizeTargetDate_(targetDateStr) {
    var now = new Date();
    var dateText = targetDateStr || Utilities.formatDate(now, TIME_ZONE, 'yyyy-MM-dd');
    var parts = dateText.split('-');

    return {
        year: parseInt(parts[0], 10),
        month: parseInt(parts[1], 10),
        day: parseInt(parts[2], 10)
    };
}

function fetchCsvRows_(gid) {
    var response = UrlFetchApp.fetch(buildCsvUrl_(gid));
    return Utilities.parseCsv(response.getContentText());
}

function buildCsvUrl_(gid) {
    return PUBLIC_SHEET_BASE + (gid || DEFAULT_GID);
}

function getSheetData(gid) {
    return {
        status: 'success',
        data: fetchCsvRows_(gid || DEFAULT_GID)
    };
}

function normalizeRecipients_(recipients) {
    var parts = String(recipients || '')
        .split(',')
        .map(function (item) {
            return item.replace(/^\s+|\s+$/g, '');
        })
        .filter(function (item) {
            return !!item;
        });
    return parts.join(',');
}

function normalizeTime_(timeString) {
    var match = String(timeString || '').match(/^(\d{1,2}):(\d{2})$/);
    var hour;
    var minute;

    if (!match) {
        return DEFAULT_TIME;
    }

    hour = Math.max(0, Math.min(23, parseInt(match[1], 10)));
    minute = Math.max(0, Math.min(59, parseInt(match[2], 10)));

    return pad2_(hour) + ':' + pad2_(minute);
}

function pad2_(value) {
    return value < 10 ? '0' + value : String(value);
}

function safeCell_(row, index) {
    if (!row || index < 0 || index >= row.length) {
        return '-';
    }
    var value = row[index];
    if (value === undefined || value === null || value === '') {
        return '-';
    }
    return String(value);
}

function escapeHtml_(value) {
    return String(value === undefined || value === null ? '' : value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function safeErrorMessage_(error) {
    return error && error.toString ? error.toString() : String(error);
}

function isMailTriggerHandler_(handlerName) {
    return (
        handlerName === 'scheduledCheckAndSend' ||
        handlerName === 'dailySendMailTask' ||
        handlerName === 'hourlyCheckAndSend' ||
        handlerName === 'scheduledSendMailTask'
    );
}
