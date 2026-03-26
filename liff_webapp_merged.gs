/**
 * ============================================================
 *  統一 Web App 入口 — 合併身體組成儀表板 + 飲食紀錄 API
 * ============================================================
 *  支援的查詢參數：
 *    ?userId=LINE_UID  → 飲食紀錄 API（patient.html 使用）
 *    ?uid=LINE_UID     → 身體組成儀表板（LIFF 模式）
 *    ?id=院區病歷號     → 身體組成（個人化連結）
 *    ?all=true         → 醫生總覽
 *    ?preheat=1        → 預熱（快速回應）
 *    ?homeWeight=院區病歷號 → 居家體重紀錄查詢
 *    POST ?saveHomeWeight=true → 儲存居家體重
 *
 *  ⚠️ 所有變數名稱皆以 DASH_ 為前綴，
 *     避免與 程式碼.gs 中的 const 重複宣告。
 * ============================================================
 */

// ===================== 🔧 設定區 =====================
var DASH_LIFF_ID = '2009326025-XAQhb1rA';
var DASH_SS_ID   = '1M2SzVfAM8I9RnuTj6_nE9Iawn9obJexo2xHl_Lyf1eQ';
// =====================================================


/**
 * ★ 統一 doGet — 根據參數分流到不同功能
 */
function doGet(e) {
  var data = { error: '請提供查詢參數（userId、uid、id 或 all）' };

  if (e && e.parameter) {
    // ── 預熱請求（快速回應，喚醒 GAS）──
    if (e.parameter.preheat) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 飲食紀錄 API（patient.html 使用）──
    if (e.parameter.userId) {
      data = _getMealRecords(String(e.parameter.userId).trim());
    }
    // ── 身體組成儀表板（LIFF 模式）──
    else if (e.parameter.uid) {
      data = dashGetDataByUid(String(e.parameter.uid).trim());
    }
    // ── 身體組成（個人化連結）──
    else if (e.parameter.id) {
      data = dashGetDataById(String(e.parameter.id).trim());
    }
    // ── 醫生總覽 ──
    else if (e.parameter.all === 'true') {
      data = dashGetAllPatients();
    }
    // ── 用院區病歷號查 LINE UID 並回傳餐食紀錄 ──
    else if (e.parameter.patientId) {
      var pSn = String(e.parameter.patientId).trim();
      var lookup = _lookupPatientSn(pSn);
      if (lookup.lineUid) {
        data = _getMealRecords(lookup.lineUid);
        data.lineUid = lookup.lineUid;
      } else {
        data = { error: '找不到此病歷號對應的 LINE UID', records: [] };
      }
    }
    // ── 飲食計畫排程查詢 ──
    else if (e.parameter.dietPlan) {
      data = _getDietPlan(String(e.parameter.dietPlan).trim());
    }
    // ── 飲食計畫歷史（看診紀錄 tab 用）──
    else if (e.parameter.dietHistory) {
      data = _getDietHistory(String(e.parameter.dietHistory).trim());
    }
    // ── 居家體重紀錄查詢 ──
    else if (e.parameter.homeWeight) {
      data = _getHomeWeight(String(e.parameter.homeWeight).trim());
    }
    // ── LINE UID → 院區病歷號 查詢 ──
    else if (e.parameter.lookupUid) {
      data = _lookupUid(String(e.parameter.lookupUid).trim());
    }
    // ── 員工白名單驗證 ──
    else if (e.parameter.checkStaff) {
      data = _checkStaff(String(e.parameter.checkStaff).trim());
    }
    // ── 儲存飲食計畫（GET 避免 CORS，含日期覆寫規則）──
    else if (e.parameter.saveDietPlan) {
      return _saveDietPlan({
        patientSn:          String(e.parameter.patientSn || '').trim(),
        proteinRec:         parseInt(e.parameter.proteinRec) || 0,
        mid:                e.parameter.mid  || '',
        high:               e.parameter.high || '',
        low:                e.parameter.low  || '',
        fast:               e.parameter.fast || '',
        schedule:           e.parameter.schedule || '[]',
        weeks:              parseInt(e.parameter.weeks) || 4,
        patientTargetWeight: parseFloat(e.parameter.patientTargetWeight) || 0,
        cardMsg:            e.parameter.cardMsg ? JSON.parse(e.parameter.cardMsg) : [],
        foodTable:          e.parameter.foodTable ? JSON.parse(e.parameter.foodTable) : []
      });
    }
    // ── 儲存居家體重（GET 避免 CORS）──
    else if (e.parameter.saveHomeWeight) {
      return _saveHomeWeight({
        patientSn: String(e.parameter.patientSn || '').trim(),
        date:      String(e.parameter.date || '').trim(),
        weight:    e.parameter.weight,
        dayType:   String(e.parameter.dayType || '').trim()
      });
    }
    // ── 同步 LINE 設定名稱到客戶清單（書籤小工具 + LIFF 頁面共用）──
    else if (e.parameter.setLineName) {
      data = _setLineName({
        sName: String(e.parameter.sName || '').trim(),
        dName: String(e.parameter.dName || '').trim(),
        uid:   String(e.parameter.setLineName || '').trim()
      });
    }
    // ── 儲存醫師建議蛋白質攝取量（取代 n8n save-protein）──
    else if (e.parameter.saveProtein) {
      data = _saveProtein({
        patientSn: String(e.parameter.saveProtein || '').trim(),
        value:     parseInt(e.parameter.value) || 0
      });
    }
    // ── 查詢病人期待目標體重（weight-calendar 用）──
    else if (e.parameter.getTargetWeight) {
      data = _getTargetWeight(String(e.parameter.getTargetWeight || '').trim());
    }
    // ── 批次同步客戶清單（從測量紀錄回填最新體重/身高/性別）──
    else if (e.parameter.syncCustomers) {
      data = _syncCustomerFromMeasurements();
    }
    // ── 儲存病人期待目標體重（K 欄）──
    else if (e.parameter.saveTargetWeight) {
      data = _saveTargetWeight({
        patientSn: String(e.parameter.saveTargetWeight || '').trim(),
        value:     parseFloat(e.parameter.value) || 0
      });
    }
    // ── 刪除測量紀錄（醫師後台用）──
    else if (e.parameter.deleteMeasurement) {
      data = _deleteMeasurement(
        String(e.parameter.deleteMeasurement || '').trim(),
        String(e.parameter.date || '').trim()
      );
    }
    // ── 刪除單筆餐食紀錄（recordId）──
    else if (e.parameter.deleteMealRecord) {
      data = _deleteMealRecord(
        String(e.parameter.deleteMealRecord || '').trim(),
        String(e.parameter.recordId || '').trim()
      );
    }
    // ── 刪除整日餐食紀錄（LINE UID + date）──
    else if (e.parameter.deleteMealsByDate) {
      data = _deleteMealsByDate(
        String(e.parameter.deleteMealsByDate || '').trim(),
        String(e.parameter.date || '').trim()
      );
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * ★ 統一 doPost — 處理 POST 請求
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // ── 儲存居家體重 ──
    if (e.parameter.saveHomeWeight) {
      return _saveHomeWeight(data);
    }

    // ── 儲存飲食計畫（POST 避免 URL 長度限制）──
    if (e.parameter.saveDietPlan) {
      // POST body 的陣列欄位需要轉成 JSON 字串給 Sheet
      var _s = function(v) { return (typeof v === 'string') ? v : JSON.stringify(v || ''); };
      return _saveDietPlan({
        patientSn:          String(data.patientSn || '').trim(),
        proteinRec:         parseInt(data.proteinRec) || 0,
        mid:                _s(data.mid),
        high:               _s(data.high),
        low:                _s(data.low),
        fast:               _s(data.fast),
        schedule:           _s(data.schedule),
        weeks:              parseInt(data.weeks) || 4,
        patientTargetWeight: parseFloat(data.patientTargetWeight) || 0,
        cardMsg:            data.cardMsg || [],
        foodTable:          data.foodTable || []
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ error: '未知的 POST 請求' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
//  ★ 飲食紀錄查詢（取代 n8n webhook/patient-history）
// ============================================================
function _getMealRecords(userId) {
  try {
    var ss = SpreadsheetApp.openById(DASH_SS_ID);

    // ── 1. 查詢病患資料（客戶清單）──
    var custSheet = ss.getSheetByName('客戶清單');
    if (!custSheet) return { error: '找不到客戶清單工作表' };
    var custData = custSheet.getDataRange().getValues();
    var custH = custData[0];

    var ci = {};
    for (var c = 0; c < custH.length; c++) {
      ci[String(custH[c]).trim()] = c;
    }

    var patient = null;
    for (var i = 1; i < custData.length; i++) {
      var row = custData[i];
      if (String(row[ci['LINE UID']] || '').trim() === userId) {
        patient = {
          name:               String(row[ci['姓名']] || ''),
          id:                 String(row[ci['院區病歷號']] || ''),
          age:                row[ci['年齡']] || '',
          height:             row[ci['身高']] || '',
          weight:             parseFloat(row[ci['最新體重'] !== undefined ? ci['最新體重'] : ci['體重']]) || 0,
          recommendedProtein: parseFloat(row[ci['蛋白質攝取量']]) || 0,
          doctorRecommendedProtein: parseFloat(row[ci['醫師建議蛋白質攝取量']]) || 0,
          settingName:        String(row[ci['設定名稱']] || ''),
        };
        break;
      }
    }

    if (!patient) {
      return { error: 'patient not found', records: [], patient: null };
    }

    // ── 2. 查詢餐食紀錄 ──
    var mealSheet = ss.getSheetByName('餐食分析紀錄');
    if (!mealSheet) return { patient: patient, records: [] };
    var mealData = mealSheet.getDataRange().getValues();
    var mealH = mealData[0];

    var mi = {};
    for (var m = 0; m < mealH.length; m++) {
      mi[String(mealH[m]).trim()] = m;
    }

    // 相容兩種欄位名稱：LINE UID（空格）或 LINE_UID（底線）
    var uidColIdx = (mi['LINE UID'] !== undefined) ? mi['LINE UID']
                  : (mi['LINE_UID'] !== undefined) ? mi['LINE_UID']
                  : -1;

    var records = [];
    if (uidColIdx === -1) {
      Logger.log('餐食分析紀錄找不到 LINE UID 或 LINE_UID 欄位，現有欄位：' + Object.keys(mi).join(', '));
      return { patient: patient, records: [], debug_columns: Object.keys(mi) };
    }

    for (var r = 1; r < mealData.length; r++) {
      var row = mealData[r];
      var rowUid = String(row[uidColIdx] || '').trim();
      if (rowUid !== userId) continue;

      records.push({
        recordId:           String(row[mi['記錄ID']] || ''),
        date:               _formatMealDate(row[mi['分析時間']]),
        imageId:            String(row[mi['LINE圖片ID']] || ''),
        foods:              String(row[mi['食物清單']] || ''),
        calories:           parseFloat(row[mi['熱量(kcal)']]) || 0,
        protein:            parseFloat(row[mi['蛋白質(g)']]) || 0,
        carbs:              parseFloat(row[mi['碳水化合物(g)']]) || 0,
        fat:                parseFloat(row[mi['脂肪(g)']]) || 0,
        fiber:              parseFloat(row[mi['纖維質(g)']]) || 0,
        netCarbs:           (parseFloat(row[mi['碳水化合物(g)']]) || 0) - (parseFloat(row[mi['纖維質(g)']]) || 0),
        sodium:             parseFloat(row[mi['鈉(mg)']]) || 0,
        nutritionScore:     row[mi['營養評分']] || '',
        aiNote:             String(row[mi['AI飲食建議']] || ''),
        nutritionistAdvice: String(row[mi['營養師建議']] || ''),
        status:             String(row[mi['審核狀態']] || ''),
      });
    }

    records.sort(function(a, b) {
      return String(b.date).localeCompare(String(a.date));
    });

    return {
      patient: patient,
      records: records,
      _debug: {
        uidColumn: uidColIdx >= 0 ? Object.keys(mi).filter(function(k){ return mi[k] === uidColIdx; })[0] : 'NOT_FOUND',
        totalRows: mealData.length - 1,
        matchedRows: records.length
      }
    };

  } catch (err) {
    return { error: err.message || String(err) };
  }
}

function _formatMealDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var y = val.getFullYear();
    var M = val.getMonth() + 1;
    var d = val.getDate();
    var h = val.getHours();
    var m = val.getMinutes();
    var s = val.getSeconds();
    var ampm = h >= 12 ? '下午' : '上午';
    var h12 = h > 12 ? h - 12 : (h === 0 ? 12 : h);
    return y + '/' + M + '/' + d + ' ' + ampm + h12 + ':' +
           String(m).padStart(2, '0') + ':' + String(s).padStart(2, '0');
  }
  return String(val);
}


// ============================================================
//  身體組成儀表板 — 原有功能
// ============================================================

function dashGetDataByUid(lineUid) {
  try {
    var ss        = SpreadsheetApp.openById(DASH_SS_ID);
    var custSheet = ss.getSheetByName('客戶清單');
    if (!custSheet) return { error: '找不到客戶清單工作表' };

    var custData = custSheet.getDataRange().getValues();
    var ch       = custData[0];
    var idxUID   = ch.indexOf('LINE UID');
    var idxSN    = ch.indexOf('院區病歷號');

    if (idxUID === -1) return { error: '客戶清單缺少「LINE UID」欄位' };
    if (idxSN  === -1) return { error: '客戶清單缺少「院區病歷號」欄位' };

    for (var i = 1; i < custData.length; i++) {
      if (String(custData[i][idxUID]).trim() === lineUid) {
        var sn = String(custData[i][idxSN]).trim();
        return dashGetDataById(sn);
      }
    }

    return { error: '查無此 LINE 帳號，請聯絡店家完成登記' };

  } catch (err) {
    return { error: '查詢失敗（UID）：' + err.message };
  }
}


function dashGetDataById(custId) {
  try {
    var ss        = SpreadsheetApp.openById(DASH_SS_ID);
    var custSheet = ss.getSheetByName('客戶清單');
    var recSheet  = ss.getSheetByName('測量紀錄');

    if (!custSheet || !recSheet) return { error: '找不到工作表，請通知店家' };

    var custData = custSheet.getDataRange().getValues();
    var ch       = custData[0];
    var idxSN    = ch.indexOf('院區病歷號');
    var idxName  = ch.indexOf('姓名');
    var idxAlias = ch.indexOf('設定名稱');
    var idxWeight = ch.indexOf('最新體重');
    if (idxWeight === -1) idxWeight = ch.indexOf('體重'); // 向下相容
    var idxH     = ch.indexOf('蛋白質攝取量');
    var idxI     = ch.indexOf('醫師建議蛋白質攝取量');
    var idxFTW   = ch.indexOf('目標體重');
    var idxPTW   = ch.indexOf('病人期待目標體重');
    // 向下相容舊欄名
    if (idxPTW === -1) idxPTW = ch.indexOf('醫師建議目標體重');
    var idxHeight = ch.indexOf('身高');
    var idxGender = ch.indexOf('性別');

    if (idxSN === -1) return { error: '客戶清單缺少「院區病歷號」欄位' };

    var customer = null;
    for (var i = 1; i < custData.length; i++) {
      if (String(custData[i][idxSN]).trim() === custId) {
        customer = {
          name:        idxName  >= 0 ? String(custData[i][idxName]).trim()  : custId,
          sn:          custId,
          settingName: idxAlias >= 0 ? String(custData[i][idxAlias]).trim() : '',
          weight:                    idxWeight >= 0 ? (parseFloat(custData[i][idxWeight]) || 0) : 0,
          recommendedProtein:        idxH >= 0 ? (parseFloat(custData[i][idxH]) || 0) : 0,
          doctorRecommendedProtein:  idxI >= 0 ? (parseFloat(custData[i][idxI]) || 0) : 0,
          targetWeight:              idxFTW >= 0 ? (parseFloat(custData[i][idxFTW]) || 0) : 0,
          patientTargetWeight:       idxPTW >= 0 ? (parseFloat(custData[i][idxPTW]) || 0) : 0,
          height:                    idxHeight >= 0 ? (parseFloat(custData[i][idxHeight]) || 0) : 0,
          gender:                    idxGender >= 0 ? String(custData[i][idxGender]).trim() : ''
        };
        break;
      }
    }

    if (!customer) return { error: '查無客戶資料（' + custId + '），請聯絡店家確認' };

    var recData      = recSheet.getDataRange().getValues();
    var headers      = recData[0];
    var measurements = [];

    for (var r = 1; r < recData.length; r++) {
      if (String(recData[r][0]).trim() === custId) {
        var row = {};
        for (var c = 0; c < headers.length; c++) {
          var val = recData[r][c];
          if (val instanceof Date) {
            val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
          }
          row[String(headers[c])] = val;
        }
        measurements.push(row);
      }
    }

    if (measurements.length === 0) return { error: '尚無測量紀錄' };

    measurements.sort(function(a, b) {
      return new Date(a['測量日期/時間']) - new Date(b['測量日期/時間']);
    });

    // ── 自動計算 J欄（目標體重）：最早測量體重 × 0.85 ──
    if (measurements.length > 0 && idxFTW >= 0) {
      var earliestW = parseFloat(measurements[0]['體重']) || 0;
      if (earliestW > 0) {
        var formulaTarget = Math.round(earliestW * 0.85 * 10) / 10;
        var currentJ = customer.targetWeight || 0;
        // 寫入或更新 J 欄（每次都以最早測量為準）
        if (currentJ !== formulaTarget) {
          // 找到客戶 row index（i 是 0-based data index, sheet row = i+1）
          for (var ri = 1; ri < custData.length; ri++) {
            if (String(custData[ri][idxSN]).trim() === custId) {
              custSheet.getRange(ri + 1, idxFTW + 1).setValue(formulaTarget);
              customer.targetWeight = formulaTarget;
              break;
            }
          }
        }
      }
    }

    // ── 自動同步最新體重、身高、性別到客戶清單 ──
    var latestM = measurements[measurements.length - 1];
    var latestWeight = parseFloat(latestM['體重']) || 0;
    var latestHeight = parseFloat(latestM['身高']) || 0;
    var latestGender = String(latestM['性別'] || '').trim();
    for (var ri = 1; ri < custData.length; ri++) {
      if (String(custData[ri][idxSN]).trim() === custId) {
        // G欄：最新體重（每次都更新）
        if (latestWeight > 0 && idxWeight >= 0) {
          var curW = parseFloat(custData[ri][idxWeight]) || 0;
          if (curW !== latestWeight) {
            custSheet.getRange(ri + 1, idxWeight + 1).setValue(latestWeight);
            customer.weight = latestWeight;
          }
        }
        // 身高：只在空白時補填
        if (latestHeight > 0 && idxHeight >= 0 && !custData[ri][idxHeight]) {
          custSheet.getRange(ri + 1, idxHeight + 1).setValue(latestHeight);
          customer.height = latestHeight;
        }
        // 性別：只在空白時補填
        if (latestGender && idxGender >= 0 && !String(custData[ri][idxGender]).trim()) {
          custSheet.getRange(ri + 1, idxGender + 1).setValue(latestGender);
          customer.gender = latestGender;
        }
        break;
      }
    }

    return { customer: customer, measurements: measurements };

  } catch (err) {
    return { error: '查詢失敗（ID）：' + err.message };
  }
}


/**
 * 批次同步：從測量紀錄回填每位客戶的最新體重、身高、性別
 * 呼叫方式：GET ?syncCustomers=true
 */
function _syncCustomerFromMeasurements() {
  try {
    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var custSheet = ss.getSheetByName('客戶清單');
    var recSheet  = ss.getSheetByName('測量紀錄');
    if (!custSheet || !recSheet) return { ok: false, error: '找不到工作表' };

    var custData = custSheet.getDataRange().getValues();
    var ch = custData[0];
    var idxSN     = ch.indexOf('院區病歷號');
    var idxWeight = ch.indexOf('最新體重');
    if (idxWeight === -1) idxWeight = ch.indexOf('體重');
    var idxHeight = ch.indexOf('身高');
    var idxGender = ch.indexOf('性別');

    // 若客戶清單缺「性別」欄，自動新增
    if (idxGender === -1) {
      var lastCol = ch.length;
      custSheet.getRange(1, lastCol + 1).setValue('性別');
      idxGender = lastCol;
    }

    // 讀取測量紀錄，找每位病人的最新一筆
    var recData = recSheet.getDataRange().getValues();
    var rh = recData[0];
    var rSN     = rh.indexOf('院區病歷號');
    var rWeight = rh.indexOf('體重');
    var rHeight = rh.indexOf('身高');
    var rGender = rh.indexOf('性別');
    var rDate   = rh.indexOf('測量日期/時間');

    var latest = {};
    for (var r = 1; r < recData.length; r++) {
      var sn = String(recData[r][rSN]).trim();
      if (!sn) continue;
      var d = recData[r][rDate];
      var ts = d instanceof Date ? d.getTime() : new Date(d).getTime();
      if (!latest[sn] || ts > latest[sn].ts) {
        latest[sn] = {
          ts: ts,
          weight: rWeight >= 0 ? (parseFloat(recData[r][rWeight]) || 0) : 0,
          height: rHeight >= 0 ? (parseFloat(recData[r][rHeight]) || 0) : 0,
          gender: rGender >= 0 ? String(recData[r][rGender]).trim() : ''
        };
      }
    }

    // 更新客戶清單
    var updated = 0;
    for (var i = 1; i < custData.length; i++) {
      var sn = String(custData[i][idxSN]).trim();
      if (!sn || !latest[sn]) continue;
      var rec = latest[sn];
      var changed = false;

      if (rec.weight > 0 && idxWeight >= 0) {
        var curW = parseFloat(custData[i][idxWeight]) || 0;
        if (curW !== rec.weight) { custSheet.getRange(i + 1, idxWeight + 1).setValue(rec.weight); changed = true; }
      }
      if (rec.height > 0 && idxHeight >= 0 && !custData[i][idxHeight]) {
        custSheet.getRange(i + 1, idxHeight + 1).setValue(rec.height); changed = true;
      }
      if (rec.gender && idxGender >= 0 && !String(custData[i][idxGender]).trim()) {
        custSheet.getRange(i + 1, idxGender + 1).setValue(rec.gender); changed = true;
      }
      if (changed) updated++;
    }

    return { ok: true, updated: updated };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}


function dashGetAllPatients() {
  try {
    var ss        = SpreadsheetApp.openById(DASH_SS_ID);
    var custSheet = ss.getSheetByName('客戶清單');
    var recSheet  = ss.getSheetByName('測量紀錄');

    if (!custSheet || !recSheet) return { error: '找不到工作表' };

    var recData    = recSheet.getDataRange().getValues();
    var recHeaders = recData[0];
    var latestMap  = {};
    var countMap   = {};
    var todayStr   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var recentList = [];

    for (var r = 1; r < recData.length; r++) {
      var sn = String(recData[r][0]).trim();
      if (!sn) continue;

      var row = {};
      for (var c = 0; c < recHeaders.length; c++) {
        var val = recData[r][c];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
        }
        row[String(recHeaders[c])] = val;
      }

      countMap[sn] = (countMap[sn] || 0) + 1;
      recentList.push(row);

      if (!latestMap[sn] ||
          new Date(row['測量日期/時間']) > new Date(latestMap[sn]['測量日期/時間'])) {
        latestMap[sn] = row;
      }
    }

    var custData = custSheet.getDataRange().getValues();
    var ch       = custData[0];
    var idxSN    = ch.indexOf('院區病歷號');
    var idxName  = ch.indexOf('姓名');
    var idxAlias = ch.indexOf('設定名稱');
    var idxUID   = ch.indexOf('LINE UID');
    var idxAge   = ch.indexOf('年齡');
    var idxGender= ch.indexOf('性別');

    var patients = [];
    for (var i = 1; i < custData.length; i++) {
      var sn = String(custData[i][idxSN]).trim();
      if (!sn) continue;
      patients.push({
        sn:          sn,
        name:        idxName   >= 0 ? String(custData[i][idxName]).trim()   : sn,
        settingName: idxAlias  >= 0 ? String(custData[i][idxAlias]).trim()  : '',
        age:         idxAge    >= 0 ? custData[i][idxAge]   : '',
        gender:      idxGender >= 0 ? String(custData[i][idxGender]).trim() : '',
        hasLineUid:  idxUID    >= 0 && !!String(custData[i][idxUID]).trim(),
        count:       countMap[sn] || 0,
        latest:      latestMap[sn] || null
      });
    }

    recentList.sort(function(a, b) {
      return new Date(b['測量日期/時間']) - new Date(a['測量日期/時間']);
    });

    var todayCount = recentList.filter(function(r) {
      return (r['測量日期/時間'] || '').startsWith(todayStr);
    }).length;

    return {
      patients:      patients,
      recentList:    recentList.slice(0, 20),
      todayCount:    todayCount,
      totalPatients: patients.length
    };

  } catch (err) {
    return { error: '醫生總覽查詢失敗：' + err.message };
  }
}


// ============================================================
//  居家體重紀錄 — 橫式格式（日期為欄位，每位病患 2 列）
//  Row 1: 院區病歷號    | 2026-03-13 | 2026-03-14 | ...
//  Row 2: A0001         | 55.1       | 55.1       | ...  (體重)
//  Row 3: A0001(碳日)   | high       | high       | ...  (碳日)
//  Row 4: B0002         | 60.3       | ...
//  Row 5: B0002(碳日)   | mid        | ...
// ============================================================

function _getHomeWeight(patientSn) {
  try {
    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var sheet = ss.getSheetByName('居家體重');

    if (!sheet) return { records: [] };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { records: [] };

    var headers = data[0]; // [院區病歷號, date1, date2, ...]
    var records = [];

    // 找到該病患的列
    var patientRow = -1;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]).trim() === patientSn) { patientRow = r; break; }
    }

    if (patientRow === -1) return { records: [] };

    // 讀取所有日期，解析 "high 55.1" 格式
    for (var c = 1; c < headers.length; c++) {
      var dateVal = headers[c];
      var dateStr = '';
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        dateStr = String(dateVal).substring(0, 10);
      }

      var cellVal = String(data[patientRow][c] || '').trim();
      if (!cellVal) continue;

      // 解析 "high 55.1" 或純數字 "55.1"
      var parts = cellVal.split(' ');
      var dayType = '';
      var weight = NaN;
      if (parts.length >= 2) {
        dayType = parts[0];
        weight = parseFloat(parts[1]);
      } else {
        weight = parseFloat(parts[0]);
      }

      if (!isNaN(weight) && weight > 0) {
        var rec = { date: dateStr, weight: weight };
        if (dayType) rec.dayType = dayType;
        records.push(rec);
      }
    }

    records.sort(function(a, b) { return a.date.localeCompare(b.date); });

    return { records: records };

  } catch (err) {
    return { error: err.message, records: [] };
  }
}


/**
 * 同步 LINE 設定名稱到客戶清單
 * 書籤模式：sName="何美卿A0002", dName="Melody" → 從 sName 提取病歷號找 row，寫入設定名稱
 * LIFF 模式：uid=LINE_UID, dName="Melody" → 用 UID 找 row，寫入設定名稱
 */
/**
 * 同步書籤：用 dName（LINE 顯示名稱）比對姓名欄找到 row，
 *           寫入 設定名稱=sName，院區病歷號=從 sName 提取
 * LIFF 頁面：用 uid（LINE UID）找到 row，寫入 設定名稱=dName
 */
function _setLineName(opts) {
  try {
    var sName = opts.sName || '';
    var dName = opts.dName || '';
    var uid   = opts.uid   || '';

    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var sheet = ss.getSheetByName('客戶清單');
    if (!sheet) return { ok: false, error: '找不到客戶清單' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idxSN    = headers.indexOf('院區病歷號');
    var idxUID   = headers.indexOf('LINE UID');
    var idxName  = headers.indexOf('姓名');
    var idxAlias = headers.indexOf('設定名稱');

    if (idxAlias === -1) return { ok: false, error: '客戶清單缺少設定名稱欄位' };

    var targetRow = -1;

    if (sName && dName) {
      // ── 書籤模式：用 dName 比對姓名欄（含模糊：去 emoji 後比對）──
      if (idxName >= 0) {
        // 先嚴格比對
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][idxName]).trim() === dName) {
            targetRow = i;
            break;
          }
        }
        // 若找不到，去除 emoji 後再比（姓名欄可能帶 emoji）
        if (targetRow === -1) {
          var stripEmoji = function(s) {
            // 移除 surrogate pairs（emoji 在 UTF-16 中是 surrogate pair）及常見符號
            return s.replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '')
                    .replace(/[\u2600-\u27BF\uFE00-\uFEFF\u200D\u20E3]/g, '')
                    .trim();
          };
          var dClean = stripEmoji(dName);
          for (var i = 1; i < data.length; i++) {
            var cellClean = stripEmoji(String(data[i][idxName]).trim());
            if (cellClean === dClean) {
              targetRow = i;
              break;
            }
          }
        }
      }
      // 備援：用 dName 比對設定名稱欄（可能之前已同步過）
      if (targetRow === -1) {
        for (var i = 1; i < data.length; i++) {
          var alias = String(data[i][idxAlias]).trim();
          if (alias && alias.indexOf(dName) >= 0) {
            targetRow = i;
            break;
          }
        }
      }

      if (targetRow === -1) {
        return { ok: false, error: '找不到姓名「' + dName + '」的用戶' };
      }

      // 寫入設定名稱
      sheet.getRange(targetRow + 1, idxAlias + 1).setValue(sName);

      // 從 sName 提取院區病歷號（如 "黃錦崧A0001" → "A0001"）並直接寫入
      if (idxSN >= 0) {
        var snMatch = sName.match(/[A-Za-z]\d+/);
        if (snMatch) {
          sheet.getRange(targetRow + 1, idxSN + 1).setValue(snMatch[0]);
        }
      }

      return { ok: true, name: dName, sName: sName };

    } else if (uid && uid !== 'true') {
      // ── LIFF 模式：用 LINE UID 找 row，寫入設定名稱 ──
      if (idxUID >= 0) {
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][idxUID]).trim() === uid) {
            targetRow = i;
            break;
          }
        }
      }
      if (targetRow === -1) return { ok: false, error: '找不到此 UID' };

      // LIFF 模式只更新設定名稱（如果目前是空的才寫，避免覆蓋工作人員設的）
      var currentAlias = String(data[targetRow][idxAlias]).trim();
      if (!currentAlias && dName) {
        sheet.getRange(targetRow + 1, idxAlias + 1).setValue(dName);
      }

      return { ok: true };

    } else {
      return { ok: false, error: '缺少必要參數' };
    }
  } catch (err) {
    return { ok: false, error: err.message };
  }
}


function _saveHomeWeight(data) {
  try {
    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var sheet = ss.getSheetByName('居家體重');

    // 分頁不存在 → 自動建立
    if (!sheet) {
      sheet = ss.insertSheet('居家體重');
      sheet.getRange(1, 1).setValue('院區病歷號');
      sheet.getRange(1, 1).setFontWeight('bold');
    }

    var patientSn = data.patientSn;
    var dateStr = data.date;
    var weight = parseFloat(data.weight);
    var dayType = data.dayType || '';

    if (!patientSn || !dateStr || isNaN(weight)) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: '缺少必要參數' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var allData = sheet.getDataRange().getValues();
    var headers = allData[0];

    // 1. 找日期欄位，沒有就新增到最右邊
    var dateCol = -1;
    for (var c = 1; c < headers.length; c++) {
      var hDate = headers[c];
      var hStr = '';
      if (hDate instanceof Date) {
        hStr = Utilities.formatDate(hDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        hStr = String(hDate).substring(0, 10);
      }
      if (hStr === dateStr) {
        dateCol = c + 1; // 1-based
        break;
      }
    }

    if (dateCol === -1) {
      dateCol = headers.length + 1;
      sheet.getRange(1, dateCol).setValue(dateStr);
    }

    // 2. 找病患列（一人一列）
    var patientRow = -1;
    for (var r = 1; r < allData.length; r++) {
      if (String(allData[r][0]).trim() === patientSn) { patientRow = r + 1; break; }
    }

    // 沒有 → 新增一列
    if (patientRow === -1) {
      patientRow = allData.length + 1;
      sheet.getRange(patientRow, 1).setValue(patientSn);
    }

    // 3. 寫入合併格式 "dayType weight"，例如 "high 55.1"
    var cellValue = dayType ? (dayType + ' ' + weight) : String(weight);
    sheet.getRange(patientRow, dateCol).setValue(cellValue);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * LINE UID → 院區病歷號 查詢（從「客戶清單」分頁）
 * A欄: LINE UID, C欄: 院區病歷號
 */
function _lookupUid(uid) {
  try {
    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var sheet = ss.getSheetByName('客戶清單');
    if (!sheet) return { error: '找不到客戶清單分頁' };

    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]).trim() === uid) {
        return { patientSn: String(data[r][2]).trim() };
      }
    }
    return { error: '找不到此 LINE 帳號對應的病歷號' };
  } catch (err) {
    return { error: err.message };
  }
}


/**
 * 院區病歷號 → LINE UID 反查（從「客戶清單」分頁）
 * A欄: LINE UID, C欄: 院區病歷號
 */
function _lookupPatientSn(patientSn) {
  try {
    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var sheet = ss.getSheetByName('客戶清單');
    if (!sheet) return { error: '找不到客戶清單分頁' };

    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][2]).trim() === patientSn) {
        return { lineUid: String(data[r][0]).trim() };
      }
    }
    return { error: '找不到此病歷號對應的 LINE UID' };
  } catch (err) {
    return { error: err.message };
  }
}


/**
 * 讀取飲食計畫排程（從「飲食計劃」分頁）
 * A: 院區病歷號, B: 日期, D: 建議蛋白質, I: 排程JSON, K: 週數
 * 回傳: { schedule: [...], date: "YYYY-MM-DD", weeks: N, protein: N }
 */
function _getDietPlan(patientSn) {
  try {
    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var sheet = ss.getSheetByName('飲食計劃');
    if (!sheet) return { error: '找不到飲食計劃分頁' };

    var data = sheet.getDataRange().getValues();
    // 找到最後一筆該病患的記錄（可能有多筆，取最新的）
    var latestRow = -1;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]).trim() === patientSn) {
        latestRow = r;
      }
    }

    if (latestRow === -1) return { error: '找不到此病患的飲食計畫' };

    var row = data[latestRow];

    // B欄: 日期
    var dateVal = row[1];
    var dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      dateStr = String(dateVal).substring(0, 10);
    }

    // I欄 (index 8): 排程 JSON
    var scheduleRaw = String(row[8] || '').trim();
    var schedule = [];
    try {
      schedule = JSON.parse(scheduleRaw);
    } catch(e) {
      schedule = ['mid','mid','mid','mid','mid','mid','mid'];
    }

    // J欄 (index 9): 週數
    var weeksFromRow = row.length > 9 ? row[9] : null;
    var weeks = parseInt(weeksFromRow) || 0;
    if (!weeks) {
      // fallback: 直接從 Sheet 讀取 J 欄
      weeks = parseInt(sheet.getRange(latestRow + 1, 10).getValue()) || 4;
    }

    // D欄 (index 3): 建議蛋白質
    var protein = parseFloat(row[3]) || 0;

    return {
      schedule: schedule,
      date: dateStr,
      weeks: weeks,
      protein: protein
    };

  } catch (err) {
    return { error: err.message };
  }
}


/**
 * 飲食計畫歷史（看診紀錄 tab 用）
 * 回傳 { records: [ { date, schedule, weeks, protein, weeklyMr, weeklyQb, totalMr, totalQb, dayPlans } ] }
 */
function _getDietHistory(patientSn) {
  try {
    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var sheet = ss.getSheetByName('飲食計劃');
    if (!sheet) return { records: [] };

    var data = sheet.getDataRange().getValues();
    var records = [];

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]).trim() !== patientSn) continue;

      var row = data[r];

      // B欄: 日期
      var dateVal = row[1];
      var dateStr = '';
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        dateStr = String(dateVal).substring(0, 10);
      }

      // I欄 (index 8): 排程
      var schedule = [];
      try { schedule = JSON.parse(String(row[8] || '')); } catch(e) {}

      // J欄 (index 9): 週數
      var weeksRaw = row.length > 9 ? row[9] : null;
      var weeks = parseInt(weeksRaw) || 0;
      if (!weeks) {
        // fallback: 直接從 Sheet 讀取 J 欄
        weeks = parseInt(sheet.getRange(r + 1, 10).getValue()) || 4;
      }

      // D欄 (index 3): 蛋白質
      var protein = parseFloat(row[3]) || 0;

      // 從排程算代餐包數
      var packs = protein <= 110 ? 2 : 3;
      var weeklyMr = 0, weeklyQb = 0;
      for (var s = 0; s < schedule.length; s++) {
        var t = schedule[s];
        if (t === 'fast') { weeklyMr += 6; weeklyQb += 2; }
        else if (t === 'low') { weeklyMr += Math.min(packs + 1, 3); weeklyQb += 1; }
        else { weeklyMr += packs; }
      }

      // E~H 欄 (index 4~7): 各碳日餐食 JSON
      var MEAL_NAMES = ['早餐','早午餐','午餐','下午茶','晚餐','宵夜'];
      var MEAL_ICONS = ['🌅','🥐','☀️','🍵','🌙','🌜'];
      var COL_LABELS = ['代餐','配方','排糖包±排油包','食物蛋白','蔬菜','Q寶','中碳水','高碳水±高油脂'];
      var COL_UNITS  = ['包','包','','克','','包','',''];
      var COL_SEPS   = ['→','→','+','→','+','→','→','→'];
      var dayTypes = ['mid','high','low','fast'];
      var dayPlans = {};

      for (var di = 0; di < dayTypes.length; di++) {
        var colIdx = 4 + di; // E=4, F=5, G=6, H=7
        var rawMeals = String(row[colIdx] || '').trim();
        if (!rawMeals) continue;

        var mealGrid = [];
        try { mealGrid = JSON.parse(rawMeals); } catch(e) { continue; }
        if (!Array.isArray(mealGrid) || mealGrid.length === 0) continue;

        var dt = dayTypes[di];
        var dpProtein = protein;
        if (dt === 'low') dpProtein = protein + 20;
        if (dt === 'fast') dpProtein = 120;
        var pk = (dt === 'fast') ? { mr: 6, qb: 2 } : (dt === 'low') ? { mr: Math.min(packs+1,3), qb: 1 } : { mr: packs, qb: 0 };

        // 把原始格子陣列轉成 { icon, name, items } 格式
        var mealsFormatted = [];
        for (var mi = 0; mi < mealGrid.length && mi < 6; mi++) {
          var mRow = mealGrid[mi];
          if (!Array.isArray(mRow)) continue;
          var parts = [];
          var partSeps = [];
          for (var fi = 0; fi < mRow.length && fi < COL_LABELS.length; fi++) {
            var v = String(mRow[fi] || '').trim();
            if (v === '-' || v === '0' || v === '' || v === '0') continue;
            var lbl = COL_LABELS[fi];
            var unit = COL_UNITS[fi];
            // 記錄此 part 前面的分隔符（用 COL_SEPS 查）
            if (parts.length > 0) partSeps.push(COL_SEPS[fi] || '→');
            // 數字欄（食物蛋白）: 顯示 label + value + unit
            if (fi === 3) { parts.push(lbl + v + unit); }
            // 選擇欄值為 '1': 顯示 label + 1 + unit
            else if (v === '1' && unit) { parts.push(lbl + '1' + unit); }
            // 其他選擇欄（適量、大量、1拳頭等）: 顯示 label + value
            else { parts.push(lbl + (v === '1' ? '' : v)); }
          }
          if (parts.length === 0) continue;
          var itemsStr = '';
          for (var pi = 0; pi < parts.length; pi++) {
            if (pi > 0) itemsStr += ' ' + partSeps[pi - 1] + ' ';
            itemsStr += parts[pi];
          }
          mealsFormatted.push({
            icon: MEAL_ICONS[mi] || '',
            name: MEAL_NAMES[mi] || '',
            items: itemsStr
          });
        }

        dayPlans[dt] = {
          protein: dpProtein,
          mr: pk.mr,
          qb: pk.qb,
          meals: mealsFormatted
        };
      }

      // K欄 (index 10): 小卡
      var cardMsg = [];
      try { var cmRaw = String(row[10] || '').trim(); if (cmRaw) cardMsg = JSON.parse(cmRaw); } catch(e) {}

      // L欄 (index 11): 食物表
      var foodTable = [];
      try { var ftRaw = String(row[11] || '').trim(); if (ftRaw) foodTable = JSON.parse(ftRaw); } catch(e) {}

      records.push({
        date: dateStr,
        schedule: schedule,
        weeks: weeks,
        protein: protein,
        weeklyMr: weeklyMr,
        weeklyQb: weeklyQb,
        totalMr: weeklyMr * weeks,
        totalQb: weeklyQb * weeks,
        dayPlans: dayPlans,
        cardMsg: cardMsg,
        foodTable: foodTable,
        isCurrent: (r === data.length - 1 || records.length === 0)
      });
    }

    return { records: records };

  } catch (err) {
    return { error: err.message, records: [] };
  }
}


/**
 * 儲存飲食計畫（含日期覆寫規則）
 * 規則：同一病患 + 同一日期 → 覆寫該列
 *       同一病患 + 不同日期 → 新增一列
 *       新病患           → 新增一列
 *
 * 欄位對應：
 * A(0):院區病歷號  B(1):日期  C(2):更新時間  D(3):建議蛋白質
 * E(4):中碳日JSON  F(5):高碳日JSON  G(6):低碳日JSON  H(7):斷食日JSON
 * I(8):排程JSON    J(9):週數  K(10):小卡JSON  L(11):食物表JSON
 */
function _saveDietPlan(data) {
  try {
    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var sheet = ss.getSheetByName('飲食計劃');
    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: '找不到飲食計劃分頁' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var patientSn = data.patientSn;
    if (!patientSn) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: '缺少 patientSn' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 今天日期（台灣時區）
    var now = new Date();
    var todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var nowStr   = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // 準備寫入的資料列
    var rowData = [
      patientSn,
      todayStr,
      nowStr,
      data.proteinRec || 0,
      data.mid  || '',
      data.high || '',
      data.low  || '',
      data.fast || '',
      data.schedule || '[]',
      data.weeks || 4,
      data.cardMsg ? JSON.stringify(data.cardMsg) : '',
      data.foodTable ? JSON.stringify(data.foodTable) : ''
    ];

    var allData = sheet.getDataRange().getValues();

    // 尋找同病患 + 同日期的列（覆寫候選）
    var overwriteRow = -1;
    for (var r = 1; r < allData.length; r++) {
      if (String(allData[r][0]).trim() !== patientSn) continue;

      // 取得該列日期
      var existDate = allData[r][1];
      var existDateStr = '';
      if (existDate instanceof Date) {
        existDateStr = Utilities.formatDate(existDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        existDateStr = String(existDate).substring(0, 10);
      }

      if (existDateStr === todayStr) {
        overwriteRow = r + 1; // 1-based for Sheet API
        break;
      }
    }

    if (overwriteRow > 0) {
      // 同日期 → 覆寫該列
      sheet.getRange(overwriteRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // 不同日期或新病患 → 新增一列到最後
      sheet.appendRow(rowData);
    }

    // ── 同步更新客戶清單 K 欄（病人期待目標體重）──
    var ptw = data.patientTargetWeight;
    if (ptw && ptw > 0) {
      try {
        var custSheet = ss.getSheetByName('客戶清單');
        if (custSheet) {
          var custData = custSheet.getDataRange().getValues();
          var custH = custData[0];
          var cIdxSN  = custH.indexOf('院區病歷號');
          var cIdxPTW = custH.indexOf('病人期待目標體重');
          if (cIdxPTW === -1) {
            cIdxPTW = custH.length;
            custSheet.getRange(1, cIdxPTW + 1).setValue('病人期待目標體重');
          }
          if (cIdxSN >= 0) {
            for (var cr = 1; cr < custData.length; cr++) {
              if (String(custData[cr][cIdxSN]).trim() === patientSn) {
                custSheet.getRange(cr + 1, cIdxPTW + 1).setValue(ptw);
                break;
              }
            }
          }
        }
      } catch(e) { /* 不影響主流程 */ }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, action: overwriteRow > 0 ? 'overwrite' : 'append', date: todayStr }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * 儲存醫師建議蛋白質攝取量（取代 n8n save-protein webhook）
 * 用院區病歷號找到客戶清單的 row，更新「醫師建議蛋白質攝取量」欄
 */
function _saveProtein(opts) {
  try {
    var patientSn = opts.patientSn;
    var value     = opts.value;
    if (!patientSn) return { ok: false, error: '缺少 patientSn' };
    if (!value || value <= 0) return { ok: false, error: '蛋白質數值無效' };

    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var sheet = ss.getSheetByName('客戶清單');
    if (!sheet) return { ok: false, error: '找不到客戶清單' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idxSN = headers.indexOf('院區病歷號');
    var idxI  = headers.indexOf('醫師建議蛋白質攝取量');

    if (idxSN === -1) return { ok: false, error: '缺少院區病歷號欄位' };
    if (idxI  === -1) return { ok: false, error: '缺少醫師建議蛋白質攝取量欄位' };

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][idxSN]).trim() === patientSn) {
        sheet.getRange(r + 1, idxI + 1).setValue(value);
        return { ok: true, patientSn: patientSn, value: value };
      }
    }

    return { ok: false, error: '找不到病歷號「' + patientSn + '」' };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/**
 * 儲存病人期待目標體重（K 欄）
 * J 欄「目標體重」由 dashGetDataById 自動從最早測量紀錄計算寫入
 * 若 K 欄不存在，自動新增
 */
function _saveTargetWeight(opts) {
  try {
    var patientSn = opts.patientSn;
    var value     = opts.value;  // 病人期待目標體重
    if (!patientSn) return { ok: false, error: '缺少 patientSn' };
    if (!value || value <= 0) return { ok: false, error: '目標體重數值無效' };

    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var sheet = ss.getSheetByName('客戶清單');
    if (!sheet) return { ok: false, error: '找不到客戶清單' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idxSN  = headers.indexOf('院區病歷號');
    var idxPTW = headers.indexOf('病人期待目標體重');

    if (idxSN === -1) return { ok: false, error: '缺少院區病歷號欄位' };

    // 若 K 欄不存在，自動新增
    if (idxPTW === -1) {
      idxPTW = headers.length;
      sheet.getRange(1, idxPTW + 1).setValue('病人期待目標體重');
      headers.push('病人期待目標體重');
    }

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][idxSN]).trim() === patientSn) {
        sheet.getRange(r + 1, idxPTW + 1).setValue(value);
        return { ok: true, patientSn: patientSn, value: value };
      }
    }

    return { ok: false, error: '找不到病歷號「' + patientSn + '」' };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}


/**
 * 查詢病人期待目標體重（輕量端點，weight-calendar 用）
 * 只從客戶清單讀一個欄位，避免拉全部體組成資料
 */
function _getTargetWeight(patientSn) {
  try {
    if (!patientSn) return { error: '缺少 patientSn' };

    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var sheet = ss.getSheetByName('客戶清單');
    if (!sheet) return { error: '找不到客戶清單' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idxSN  = headers.indexOf('院區病歷號');
    var idxPTW = headers.indexOf('病人期待目標體重');
    if (idxPTW === -1) idxPTW = headers.indexOf('醫師建議目標體重');

    if (idxSN === -1) return { error: '缺少院區病歷號欄位' };

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][idxSN]).trim() === patientSn) {
        var val = idxPTW >= 0 ? (parseFloat(data[r][idxPTW]) || 0) : 0;
        return { ok: true, patientTargetWeight: val };
      }
    }

    return { ok: true, patientTargetWeight: 0 };
  } catch (err) {
    return { error: err.message };
  }
}


/**
 * 員工白名單驗證
 * 在「員工白名單」分頁檢查 LINE UID 是否有權限
 * 分頁結構：A欄=LINE UID, B欄=姓名, C欄=角色(doctor/staff)
 * 若分頁不存在，自動建立並把第一位登入者設為 doctor
 */
function _checkStaff(lineUid) {
  try {
    if (!lineUid) return { authorized: false, error: '缺少 LINE UID' };

    var ss = SpreadsheetApp.openById(DASH_SS_ID);

    // 只查「醫師清單」— 只有醫師才能登入後台
    // （工作人員清單僅用於 LINE Chat 自動登入，不可開後台）
    var docSheet = ss.getSheetByName('醫師清單');
    if (docSheet && docSheet.getLastRow() > 1) {
      var docData = docSheet.getDataRange().getValues();
      for (var d = 1; d < docData.length; d++) {
        if (String(docData[d][0]).trim() === lineUid) {
          return {
            authorized: true,
            role: 'doctor',
            name: String(docData[d][1] || '').trim()
          };
        }
      }
    }

    return { authorized: false };
  } catch (err) {
    return { authorized: false, error: err.message };
  }
}


/**
 * 刪除測量紀錄（醫師後台用）
 * 根據院區病歷號 + 日期找到該列並刪除
 */
function _deleteMeasurement(patientSn, dateStr) {
  try {
    if (!patientSn || !dateStr) return { error: '缺少病歷號或日期' };
    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var recSheet = ss.getSheetByName('測量紀錄');
    if (!recSheet) return { error: '找不到測量紀錄工作表' };

    var data = recSheet.getDataRange().getValues();
    var headers = data[0];
    var dateIdx = headers.indexOf('測量日期/時間');
    if (dateIdx === -1) return { error: '找不到測量日期欄位' };

    // 從底部往上找，避免刪除後 row index 跑掉
    var deleted = 0;
    for (var r = data.length - 1; r >= 1; r--) {
      if (String(data[r][0]).trim() !== patientSn) continue;
      var cellDate = data[r][dateIdx];
      var formatted = '';
      if (cellDate instanceof Date) {
        formatted = Utilities.formatDate(cellDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        formatted = String(cellDate).substring(0, 10);
      }
      if (formatted === dateStr) {
        recSheet.deleteRow(r + 1); // sheet row 是 1-based
        deleted++;
      }
    }

    if (deleted === 0) return { error: '找不到 ' + patientSn + ' 在 ' + dateStr + ' 的紀錄' };
    return { ok: true, deleted: deleted };
  } catch (err) {
    return { error: err.message };
  }
}


/**
 * 刪除單筆餐食紀錄（by LINE UID + recordId）
 * ?deleteMealRecord=<lineUid>&recordId=<recordId>
 */
function _deleteMealRecord(lineUid, recordId) {
  try {
    if (!lineUid || !recordId) return { error: '缺少 LINE UID 或記錄ID' };
    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var sheet = ss.getSheetByName('餐食分析紀錄');
    if (!sheet) return { error: '找不到餐食分析紀錄工作表' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var mi = {};
    for (var c = 0; c < headers.length; c++) mi[String(headers[c]).trim()] = c;

    var uidColIdx = (mi['LINE UID'] !== undefined) ? mi['LINE UID']
                  : (mi['LINE_UID'] !== undefined) ? mi['LINE_UID']
                  : -1;
    var ridColIdx = (mi['記錄ID'] !== undefined) ? mi['記錄ID'] : -1;
    if (uidColIdx === -1 || ridColIdx === -1) return { error: '找不到必要欄位' };

    var deleted = 0;
    for (var r = data.length - 1; r >= 1; r--) {
      var rowUid = String(data[r][uidColIdx] || '').trim();
      var rowRid = String(data[r][ridColIdx] || '').trim();
      if (rowUid === lineUid && rowRid === recordId) {
        sheet.deleteRow(r + 1);
        deleted++;
      }
    }

    if (deleted === 0) return { error: '找不到此筆紀錄' };
    return { ok: true, deleted: deleted };
  } catch (err) {
    return { error: err.message };
  }
}


/**
 * 刪除整日餐食紀錄（by LINE UID + date key like "2026-03-20"）
 * ?deleteMealsByDate=<lineUid>&date=<dateKey>
 */
function _deleteMealsByDate(lineUid, dateKey) {
  try {
    if (!lineUid || !dateKey) return { error: '缺少 LINE UID 或日期' };
    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var sheet = ss.getSheetByName('餐食分析紀錄');
    if (!sheet) return { error: '找不到餐食分析紀錄工作表' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var mi = {};
    for (var c = 0; c < headers.length; c++) mi[String(headers[c]).trim()] = c;

    var uidColIdx = (mi['LINE UID'] !== undefined) ? mi['LINE UID']
                  : (mi['LINE_UID'] !== undefined) ? mi['LINE_UID']
                  : -1;
    var dateColIdx = (mi['分析時間'] !== undefined) ? mi['分析時間'] : -1;
    if (uidColIdx === -1 || dateColIdx === -1) return { error: '找不到必要欄位' };

    // dateKey 格式 "2026-03-20"，需與分析時間的日期部分比對
    var deleted = 0;
    for (var r = data.length - 1; r >= 1; r--) {
      var rowUid = String(data[r][uidColIdx] || '').trim();
      if (rowUid !== lineUid) continue;

      var cellDate = data[r][dateColIdx];
      var formatted = '';
      if (cellDate instanceof Date) {
        formatted = Utilities.formatDate(cellDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        // 格式如 "2026/3/20 12:30" → 轉成 "2026-03-20"
        var m = String(cellDate).match(/(\d{4})\/(\d+)\/(\d+)/);
        if (m) {
          formatted = m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
        } else {
          formatted = String(cellDate).substring(0, 10);
        }
      }

      if (formatted === dateKey) {
        sheet.deleteRow(r + 1);
        deleted++;
      }
    }

    if (deleted === 0) return { error: '找不到 ' + dateKey + ' 的紀錄' };
    return { ok: true, deleted: deleted };
  } catch (err) {
    return { error: err.message };
  }
}
