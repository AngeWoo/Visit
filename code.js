var TIME_ZONE = 'Asia/Taipei';
var DEFAULT_GID = '1401484943';
var PUBLIC_SHEET_BASE =
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vSG6BO-kJ_GZoqn-JbMLhC_mDmlJ-q_5eWL4gUFnpoIrfvFf0iJ2uk4r0eQGZ9sfFVqL5Dx_UrEVOjI/pub?output=csv&gid=';
var DEFAULT_RECIPIENTS = 'ange.wu@ycgroup.tw';
var DEFAULT_TIME = '16:50';
var MAIL_SENDER_NAME = '\u5404\u5ee0\u5340\u8a2a\u5ba2\u7cfb\u7d71';
var MAIL_SUBJECT_PREFIX = '\u5404\u5ee0\u5340\u8a2a\u5ba2\u65e5\u5831';
var LABEL_NOT_SENT = '\u5c1a\u672a\u5bc4\u51fa';
var SHEETS_INFO = [
    { name: '\u5167\u6e56', gid: '1401484943' },
    { name: '\u694a\u6885', gid: '930740199' },
    { name: '\u5f70\u6ff1\u8584\u819c', gid: '412432769' },
    { name: '\u5f70\u6ff1\u81a0\u5e36', gid: '545698913' }
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
            message: safeErrorMessage_(err)
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
    var i;

    try {
        triggers = ScriptApp.getProjectTriggers();
    } catch (e) {
        triggers = [];
    }

    for (i = 0; i < triggers.length; i++) {
        if (isMailTriggerHandler_(triggers[i].getHandlerFunction())) {
            triggerCount++;
        }
    }

    return {
        recipients: props.getProperty('MAIL_RECIPIENTS') || DEFAULT_RECIPIENTS,
        time: props.getProperty('MAIL_TIME') || DEFAULT_TIME,
        triggerActive: triggerCount > 0,
        triggerCount: triggerCount,
        lastSent: props.getProperty('LAST_SENT_DATETIME') || props.getProperty('LAST_SENT_DATE') || LABEL_NOT_SENT,
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
        var triggerMsg = ensureIntervalTrigger();
        var immediateCheckResult;

        props.setProperty('MAIL_RECIPIENTS', normalizedRecipients || DEFAULT_RECIPIENTS);
        props.setProperty('MAIL_TIME', normalizedTime);
        immediateCheckResult = scheduledCheckAndSend();

        return {
            success: true,
            message:
                '\u8a2d\u5b9a\u5df2\u5132\u5b58\uff0c\u6536\u4ef6\u4eba\uff1a' +
                (normalizedRecipients || DEFAULT_RECIPIENTS) +
                '\uff0c\u6392\u7a0b\u6642\u9593\uff1a' +
                normalizedTime +
                '\u3002' +
                triggerMsg +
                '\u7acb\u5373\u6aa2\u67e5\u7d50\u679c\uff1a' +
                formatImmediateCheckMessage_(immediateCheckResult)
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

    return {
        success: true,
        message: '\u5bc4\u4fe1\u72c0\u614b\u8207\u932f\u8aa4\u7d00\u9304\u5df2\u6e05\u9664\u3002'
    };
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

        return '\u5df2\u5efa\u7acb\u6bcf 10 \u5206\u9418\u6aa2\u67e5\u4e00\u6b21\u7684\u6392\u7a0b\u89f8\u767c\u5668\u3002';
    } catch (e) {
        PropertiesService.getScriptProperties().setProperty(
            'LAST_ERROR',
            '\u5efa\u7acb\u89f8\u767c\u5668\u5931\u6557\uff1a' + safeErrorMessage_(e)
        );
        return '\u5efa\u7acb\u89f8\u767c\u5668\u5931\u6557\uff1a' + safeErrorMessage_(e);
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
            return { success: false, code: 'weekend', message: '\u4eca\u5929\u662f\u9031\u672b\uff0c\u4e0d\u57f7\u884c\u6392\u7a0b\u5bc4\u4fe1\u3002' };
        }

        conf = getMailConfig();
        if (!conf.recipients || !conf.time) {
            return { success: false, code: 'missing_config', message: '\u6536\u4ef6\u4eba\u6216\u6392\u7a0b\u6642\u9593\u672a\u8a2d\u5b9a\u5b8c\u6574\u3002' };
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
            return {
                success: false,
                code: 'outside_window',
                message:
                    '\u76ee\u524d\u6642\u9593\u4e0d\u5728\u53ef\u5bc4\u9001\u8996\u7a97\u5167\uff0c\u9700\u8981\u5728 ' +
                    conf.time +
                    ' \u5f8c 10 \u5206\u9418\u5167\u624d\u6703\u5bc4\u9001\u3002'
            };
        }

        todayStr = Utilities.formatDate(now, TIME_ZONE, 'yyyy-MM-dd');
        scheduleSlotKey = todayStr + ' ' + conf.time;
        lastScheduledSlot = props.getProperty('LAST_SCHEDULED_SENT_KEY') || '';
        if (lastScheduledSlot === scheduleSlotKey) {
            return {
                success: false,
                code: 'already_sent',
                message: '\u4eca\u5929\u9019\u500b\u6392\u7a0b\u6642\u6bb5\u5df2\u7d93\u5bc4\u904e\u4e00\u6b21\u4e86\u3002'
            };
        }

        remaining = MailApp.getRemainingDailyQuota();
        if (remaining <= 0) {
            props.setProperty(
                'LAST_ERROR',
                nowTimeStr + ' - MailApp \u6bcf\u65e5\u914d\u984d\u5df2\u7528\u5b8c'
            );
            return {
                success: false,
                code: 'quota_exhausted',
                message: 'MailApp \u6bcf\u65e5\u914d\u984d\u5df2\u7528\u5b8c\u3002'
            };
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
        Logger.log('\u6392\u7a0b\u90f5\u4ef6\u5bc4\u9001\u5b8c\u6210\uff1a' + nowTimeStr);
        return {
            success: true,
            code: 'sent',
            message: '\u7acb\u5373\u6aa2\u67e5\u5df2\u6210\u529f\u5bc4\u51fa\u4e00\u5c01\u6392\u7a0b\u90f5\u4ef6\u3002'
        };
    } catch (e) {
        props.setProperty('LAST_ERROR', nowTimeStr + ' - ' + safeErrorMessage_(e));
        Logger.log('\u6392\u7a0b\u5bc4\u4fe1\u5931\u6557\uff1a' + safeErrorMessage_(e));
        return {
            success: false,
            code: 'error',
            message: safeErrorMessage_(e)
        };
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
            return {
                success: false,
                error: '\u8acb\u5148\u8a2d\u5b9a\u81f3\u5c11\u4e00\u4f4d\u6536\u4ef6\u4eba\u3002'
            };
        }

        if (remaining <= 0) {
            return {
                success: false,
                error:
                    'MailApp \u6bcf\u65e5\u914d\u984d\u5df2\u7528\u5b8c\uff08\u5269\u9918\uff1a' +
                    remaining +
                    '\uff09\u3002'
            };
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
            message:
                '\u5df2\u5f37\u5236\u5bc4\u51fa\u90f5\u4ef6\u7d66\uff1a' +
                targetRecipients +
                '\u3002\u5269\u9918\u914d\u984d\uff1a\u7d04 ' +
                (remaining - 1)
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
            return {
                success: false,
                error: '\u8acb\u5148\u8a2d\u5b9a\u81f3\u5c11\u4e00\u4f4d\u6536\u4ef6\u4eba\u3002'
            };
        }

        MailApp.sendEmail({
            to: targetRecipients,
            name: MAIL_SENDER_NAME,
            subject: '\u6e2c\u8a66\u90f5\u4ef6 - ' + MAIL_SUBJECT_PREFIX + ' - ' + dateStr,
            htmlBody: htmlBody
        });

        PropertiesService.getScriptProperties().setProperty('LAST_SENT_DATETIME', nowTimeStr);

        return {
            success: true,
            message: '\u6e2c\u8a66\u90f5\u4ef6\u5df2\u5bc4\u51fa\u7d66\uff1a' + targetRecipients
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
        '<p style="margin-top:24px;color:#64748b;font-size:12px;">' +
        escapeHtml_('\u7cfb\u7d71\u7522\u751f\u6642\u9593\uff1a' + nowText) +
        '</p>';
    message += '</div>';

    return message;
}

function buildFactorySectionHtml_(sheetInfo, dateInfo) {
    var rows;
    var headers = [
        '#',
        '\u6642\u9593',
        '\u59d3\u540d',
        '\u624b\u6a5f',
        '\u8a2a\u5ba2\u516c\u53f8',
        '\u62dc\u8a2a\u516c\u53f8',
        '\u62dc\u8a2a\u55ae\u4f4d',
        '\u53d7\u8a2a\u8005',
        '\u4e8b\u7531',
        '\u96e2\u5834\u6642\u9593'
    ];
    var html = '';
    var i;

    try {
        rows = normalizeFactoryRows_(sheetInfo.gid, dateInfo);
    } catch (e) {
        rows = [];
    }

    html +=
        '<h3 style="background:#e8f0fe;padding:10px 14px;border-left:5px solid #004e92;margin:20px 0 8px;font-size:15px;">' +
        escapeHtml_(sheetInfo.name) +
        ' | ' +
        escapeHtml_('\u4eca\u65e5\u7b46\u6578\uff1a' + rows.length) +
        '</h3>';

    if (!rows.length) {
        html +=
            '<p style="margin:0 0 16px 14px;color:#94a3b8;font-size:13px;">' +
            escapeHtml_('\u672c\u65e5\u7121\u7b26\u5408\u8cc7\u6599\u3002') +
            '</p>';
        return html;
    }

    html += '<table style="border-collapse:collapse;width:100%;min-width:1100px;font-size:12px;margin-bottom:16px;">';
    html += '<thead><tr style="background:#004e92;color:#ffffff;">';
    for (i = 0; i < headers.length; i++) {
        html +=
            '<th style="padding:8px;border:1px solid #dbeafe;text-align:left;white-space:nowrap;">' +
            escapeHtml_(headers[i]) +
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
    var idx = detectColumnIndexes_(header);
    var normalizedRows = [];
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
            ['\u59d3\u540d'],
            ['\u8a2a\u5ba2', '\u59d3\u540d'],
            ['\u4f86\u8a2a', '\u59d3\u540d']
        ]),
        phone: findHeaderIndex_(header, [
            ['\u624b\u6a5f'],
            ['\u96fb\u8a71']
        ]),
        visitorCompany: findHeaderIndex_(header, [
            ['\u60a8\u7684\u516c\u53f8'],
            ['\u516c\u53f8\u540d\u7a31'],
            ['\u8a2a\u5ba2', '\u516c\u53f8'],
            ['\u4f86\u8a2a', '\u516c\u53f8']
        ]),
        targetCompany: findHeaderIndex_(header, [
            ['\u6b32\u62dc\u8a2a\u516c\u53f8'],
            ['\u62dc\u8a2a\u516c\u53f8']
        ]),
        targetUnit: findHeaderIndex_(header, [
            ['\u6b32\u62dc\u8a2a\u55ae\u4f4d'],
            ['\u62dc\u8a2a\u55ae\u4f4d'],
            ['\u90e8\u9580']
        ]),
        targetPerson: findHeaderIndex_(header, [
            ['\u62dc\u8a2a\u5c0d\u8c61'],
            ['\u88ab\u8a2a\u8005'],
            ['\u53d7\u8a2a\u8005'],
            ['\u63a5\u5f85\u4eba']
        ]),
        reason: findHeaderIndex_(header, [
            ['\u4e8b\u7531'],
            ['\u539f\u56e0'],
            ['\u76ee\u7684']
        ]),
        leaveTime: findHeaderIndex_(header, [
            ['\u96e2\u5834'],
            ['\u96e2\u5ee0'],
            ['\u96e2\u958b\u6642\u9593']
        ])
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
    var value;

    if (!row || index < 0 || index >= row.length) {
        return '-';
    }

    value = row[index];
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

function formatImmediateCheckMessage_(result) {
    if (!result || !result.message) {
        return '\u672a\u53d6\u5f97\u6aa2\u67e5\u7d50\u679c\u3002';
    }
    return result.message;
}

function isMailTriggerHandler_(handlerName) {
    return (
        handlerName === 'scheduledCheckAndSend' ||
        handlerName === 'dailySendMailTask' ||
        handlerName === 'hourlyCheckAndSend' ||
        handlerName === 'scheduledSendMailTask'
    );
}
