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
    // ── 飲食紀錄 API（用院區病歷號查詢，patient-view.html 使用）──
    else if (e.parameter.patientId) {
      data = _getMealRecordsByPatientId(String(e.parameter.patientId).trim());
    }
    // ── 身體組成儀表板（LIFF 模式）──
    else if (e.parameter.uid) {
      data = dashGetDataByUid(String(e.parameter.uid).trim());
    }
    // ── 身體組成（個人化連結）──
    else if (e.parameter.id) {
      data = dashGetDataById(String(e.parameter.id).trim());
    }
    // ── 看診紀錄 API（visit-record.html 使用）──
    else if (e.parameter.visitUid) {
      data = _getVisitRecords(String(e.parameter.visitUid).trim());
    }
    // ── 醫生總覽 ──
    else if (e.parameter.all === 'true') {
      data = dashGetAllPatients();
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
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
          weight:             parseFloat(row[ci['體重']]) || 0,
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
//  飲食紀錄查詢（用院區病歷號 → 找 LINE UID → 查餐食）
// ============================================================
function _getMealRecordsByPatientId(patientId) {
  try {
    var ss = SpreadsheetApp.openById(DASH_SS_ID);
    var custSheet = ss.getSheetByName('客戶清單');
    if (!custSheet) return { error: '找不到客戶清單工作表' };
    var custData = custSheet.getDataRange().getValues();
    var custH = custData[0];
    var ci = {};
    for (var c = 0; c < custH.length; c++) ci[String(custH[c]).trim()] = c;

    for (var i = 1; i < custData.length; i++) {
      if (String(custData[i][ci['院區病歷號']] || '').trim() === patientId) {
        var uid = String(custData[i][ci['LINE UID']] || '').trim();
        if (uid) {
          var result = _getMealRecords(uid);
          result.lineUid = uid;  // 回傳 UID 給前端快取
          return result;
        }
        // 無 LINE UID — 回傳病患基本資料但無餐食
        return {
          patient: {
            name: String(custData[i][ci['姓名']] || ''),
            id: patientId,
            settingName: String(custData[i][ci['設定名稱']] || ''),
            weight: parseFloat(custData[i][ci['體重']]) || 0,
            recommendedProtein: parseFloat(custData[i][ci['蛋白質攝取量']]) || 0,
            doctorRecommendedProtein: parseFloat(custData[i][ci['醫師建議蛋白質攝取量']]) || 0,
          },
          records: [],
          lineUid: '',
          noLine: true
        };
      }
    }
    return { error: '查無此病歷號（' + patientId + '）', records: [], patient: null };
  } catch (err) {
    return { error: err.message || String(err) };
  }
}


// ============================================================
//  身體組成儀表板 — 原有功能（以下不變）
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
    var idxWeight = ch.indexOf('體重');
    var idxH     = ch.indexOf('蛋白質攝取量');
    var idxI     = ch.indexOf('醫師建議蛋白質攝取量');

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
          doctorRecommendedProtein:  idxI >= 0 ? (parseFloat(custData[i][idxI]) || 0) : 0
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

    return { customer: customer, measurements: measurements };

  } catch (err) {
    return { error: '查詢失敗（ID）：' + err.message };
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
//  看診紀錄查詢（visit-record.html 使用）
// ============================================================
function _getVisitRecords(lineUid) {
  try {
    var ss = SpreadsheetApp.openById(DASH_SS_ID);

    // 查病患姓名
    var custSheet = ss.getSheetByName('客戶清單');
    var custData = custSheet.getDataRange().getValues();
    var custH = custData[0];
    var ci = {};
    for (var c = 0; c < custH.length; c++) ci[String(custH[c]).trim()] = c;

    var patientName = '';
    var patientSn = '';
    for (var i = 1; i < custData.length; i++) {
      if (String(custData[i][ci['LINE UID']] || '').trim() === lineUid) {
        patientName = String(custData[i][ci['姓名']] || '');
        patientSn = String(custData[i][ci['院區病歷號']] || '');
        break;
      }
    }

    if (!patientSn) return { error: '查無此帳號', records: [] };

    // 查看診紀錄
    var visitSheet = ss.getSheetByName('看診紀錄');
    if (!visitSheet) return { patientName: patientName, records: [] };

    var vData = visitSheet.getDataRange().getValues();
    var vh = vData[0];
    var vi = {};
    for (var v = 0; v < vh.length; v++) vi[String(vh[v]).trim()] = v;

    var records = [];
    for (var r = 1; r < vData.length; r++) {
      var row = vData[r];
      if (String(row[vi['院區病歷號']] || '').trim() !== patientSn &&
          String(row[vi['LINE UID']] || '').trim() !== lineUid) continue;

      var dateVal = row[vi['日期時間']];
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      }

      records.push({
        date:    String(dateVal || ''),
        protein: parseFloat(row[vi['蛋白質建議(g)']]) || 0,
        note:    String(row[vi['建議內容']] || ''),
        doctor:  String(row[vi['醫師姓名']] || '')
      });
    }

    // 最新的在前面
    records.sort(function(a, b) { return String(b.date).localeCompare(String(a.date)); });

    return { patientName: patientName, records: records };

  } catch (err) {
    return { error: err.message || String(err) };
  }
}
