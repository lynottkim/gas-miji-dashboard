/**
 * =================================================================
 * 미지세계 이카운트 재고조회 시스템  (v8.1)
 * -----------------------------------------------------------------
 * v8.1 변경사항
 *   - IP 기반 접근 제한 제거 (구글 계정 기반으로 대체)
 *   - 헤더에 로그인 구글 계정 표시 기능 추가
 * =================================================================
 *
 *  ┌─────────────────────────────────────────────────────────────┐
 *  │  [최초 1회 설정]  코드에 키를 직접 쓰지 마세요.              │
 *  │  1) 아래 setupCredentials() 의 값을 실제 값으로 채우세요.   │
 *  │  2) 함수를 한 번 실행하면 Script Properties 에 저장됩니다.  │
 *  │  3) 실행 후 값을 다시 XXXX 로 지워두면 코드에 키가 없어집니다. │
 *  └─────────────────────────────────────────────────────────────┘
 */

// ── 일반 설정 ────────────────────────────────────────────────────
const CONFIG = {
  ECOUNT: {
    ZONE: "CC",
    LOGIN_URL:     "https://oapiCC.ecount.com/OAPI/V2/OAPILogin",
    INVENTORY_URL: "https://oapiCC.ecount.com/OAPI/V2/InventoryBalance/GetListInventoryBalanceStatusByLocation"
  },
  SPREADSHEET: {
    ID:                "1KUx1XUcyGa5UqoPATi_stz4EmMDPfhj6eHYjY0iqdL8",
    PRESET_SHEET_NAME: "PRESET",
    TARGET_SHEET_NAME: "WH500"
  },
  SYNC_LOCK_MINUTES: 30,
  VERSION: "8.7"
};

// ══════════════════════════════════════════════════════════════════
//  [최초 1회 실행] 민감정보 → Script Properties 저장
// ══════════════════════════════════════════════════════════════════
function setupCredentials() {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    "ECOUNT_COM_CODE":     "XXXX_회사코드_XXXX",
    "ECOUNT_USER_ID":      "XXXX_사용자ID_XXXX",
    "ECOUNT_API_CERT_KEY": "XXXX_API_CERT_KEY_XXXX"
  });
  Logger.log("✅ 완료. 이 함수 안의 값들을 다시 XXXX 로 지워두세요.");
}

/** Script Properties 에서 민감정보를 읽어옵니다 */
function getCredentials_() {
  const p = PropertiesService.getScriptProperties();
  return {
    COM_CODE: p.getProperty("ECOUNT_COM_CODE")     || "",
    USER_ID:  p.getProperty("ECOUNT_USER_ID")      || "",
    CERT_KEY: p.getProperty("ECOUNT_API_CERT_KEY") || ""
  };
}

// ══════════════════════════════════════════════════════════════════
//  웹앱 진입점
// ══════════════════════════════════════════════════════════════════
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('미지세계 재고조회')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/** 시스템 정보 + 현재 로그인 계정 반환 */
function getSystemInfo() {
  const s = getLastSyncTimeInfo();
  let email = "";
  try {
    email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || "";
  } catch(e) { email = ""; }

  return {
    version:      CONFIG.VERSION,
    lastSyncTime: s.timeStr,
    lockMinutes:  CONFIG.SYNC_LOCK_MINUTES,
    userEmail:    email,
    sheetUrl:     "https://docs.google.com/spreadsheets/d/" + CONFIG.SPREADSHEET.ID
  };
}

// ══════════════════════════════════════════════════════════════════
//  보안 유틸 (민감 필드 마스킹)
// ══════════════════════════════════════════════════════════════════
function maskSecret_(v) {
  if (!v) return "";
  const s = String(v);
  if (s.length <= 6) return "***";
  return s.slice(0, 3) + "*".repeat(s.length - 6) + s.slice(-3);
}

function deepMask_(obj) {
  const KEYS = ["API_CERT_KEY","CERT_KEY","SESSION_ID","PASSWORD","PASSWD","COM_CODE"];
  function walk(node) {
    if (Array.isArray(node)) return node.map(walk);
    if (node && typeof node === "object") {
      const out = {};
      Object.keys(node).forEach(k => {
        out[k] = KEYS.includes(k.toUpperCase()) ? maskSecret_(node[k]) : walk(node[k]);
      });
      return out;
    }
    return node;
  }
  return walk(obj);
}

// ══════════════════════════════════════════════════════════════════
//  동기화 시각 파싱
// ══════════════════════════════════════════════════════════════════
function getLastSyncTimeInfo() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET.ID);
    const sheet = ss.getSheetByName(CONFIG.SPREADSHEET.TARGET_SHEET_NAME);
    if (!sheet) return { timeStr: "기록 없음", diffMins: 999 };

    const raw = sheet.getRange(1, 2).getValue();
    if (!raw) return { timeStr: "기록 없음", diffMins: 999 };

    const cellStr = raw.toString();
    const timeStr = cellStr.includes(": ") ? cellStr.split(": ").slice(1).join(": ").trim() : cellStr.trim();
    const p = timeStr.split(/[ \-:]/).map(Number);
    if (p.length < 6 || p.some(isNaN)) return { timeStr: timeStr, diffMins: 999 };

    const lastDate = new Date(p[0], p[1]-1, p[2], p[3], p[4], p[5]);
    const diffMins = Math.floor((Date.now() - lastDate.getTime()) / 60000);
    return { timeStr: timeStr, diffMins: diffMins };
  } catch(e) {
    return { timeStr: "체크 오류", diffMins: 999 };
  }
}

// ══════════════════════════════════════════════════════════════════
//  프리셋 로드
// ══════════════════════════════════════════════════════════════════
function getSeriesPresets() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET.ID);
    const sheet = ss.getSheetByName(CONFIG.SPREADSHEET.PRESET_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return [];

    return sheet.getRange(2, 1, sheet.getLastRow()-1, 9).getValues()
      .map(row => {
        const name  = (row[0] || "").toString().trim();
        const codes = [];
        for (let i=1; i<=8; i++) {
          const v = (row[i] || "").toString().trim();
          if (v) codes.push(v);
        }
        return { name, codes: codes.join(" ") };
      })
      .filter(item => item.name && item.codes);
  } catch(e) { return []; }
}

// ══════════════════════════════════════════════════════════════════
//  [동기화 엔진] 이카운트 → 구글시트 덤프
// ══════════════════════════════════════════════════════════════════
function runInventorySynchronization(forceFlag, currentWhCd) {
  const targetWh = (currentWhCd || CONFIG.SPREADSHEET.TARGET_SHEET_NAME).trim();
  const logs = [];

  if (!forceFlag) {
    const t = getLastSyncTimeInfo();
    if (t.timeStr !== "기록 없음" && t.diffMins < CONFIG.SYNC_LOCK_MINUTES) {
      return { success:false, isTimeLocked:true,
               remainingMins: CONFIG.SYNC_LOCK_MINUTES - t.diffMins,
               lastTime: t.timeStr, logs: [] };
    }
  }

  const cred = getCredentials_();
  if (!cred.CERT_KEY || cred.CERT_KEY.includes("XXXX")) {
    return { success:false, message:"API 인증키가 설정되지 않았습니다. setupCredentials()를 먼저 실행하세요.", logs:[] };
  }

  // ── 로그인 ──
  const loginPayload = {
    "COM_CODE": cred.COM_CODE, "USER_ID": cred.USER_ID,
    "API_CERT_KEY": cred.CERT_KEY, "LAN_TYPE": "ko-KR", "ZONE": CONFIG.ECOUNT.ZONE
  };
  const loginLog = { isLogin:true, requestPayload: deepMask_(loginPayload), responseRaw:null };

  let sessionId = "";
  try {
    const res    = UrlFetchApp.fetch(CONFIG.ECOUNT.LOGIN_URL, {
      method:"post", contentType:"application/json",
      payload: JSON.stringify(loginPayload), muteHttpExceptions:true
    });
    const result = JSON.parse(res.getContentText());
    loginLog.responseRaw = deepMask_(result);
    logs.push(loginLog);

    if (result.Status == "200") {
      sessionId = result.Data.Datas.SESSION_ID;
    } else {
      return { success:false, message:"이카운트 인증 실패", logs };
    }
  } catch(e) {
    loginLog.responseRaw = { error: e.toString() };
    logs.push(loginLog);
    return { success:false, message:"로그인 통신 오류", logs };
  }

  // ── 재고 덤프 ──
  const todayStr     = Utilities.formatDate(new Date(), "GMT+9", "yyyyMMdd");
  const nowTimestamp = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss");
  const payload      = { "PROD_CD":"", "WH_CD": targetWh, "BASE_DATE": todayStr };
  const apiLog       = { isLogin:false, requestPayload: payload, responseRaw:null };

  try {
    const res      = UrlFetchApp.fetch(
      `${CONFIG.ECOUNT.INVENTORY_URL}?SESSION_ID=${sessionId}`,
      { method:"post", contentType:"application/json",
        payload: JSON.stringify(payload), muteHttpExceptions:true }
    );
    const fullJson = JSON.parse(res.getContentText());
    apiLog.responseRaw = deepMask_({
      Status:      fullJson.Status,
      Message:     fullJson.Message,
      ResultCount: (fullJson.Data && fullJson.Data.Result) ? fullJson.Data.Result.length : 0
    });
    logs.push(apiLog);

    if (!fullJson || fullJson.Status != "200") {
      return { success:false, message:"재고 데이터 수신 실패", logs };
    }

    const rawItems = (fullJson.Data && fullJson.Data.Result) || [];
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET.ID);
    const sheet = ss.getSheetByName(CONFIG.SPREADSHEET.TARGET_SHEET_NAME);
    if (!sheet) {
      return { success:false, message:`시트에 '${CONFIG.SPREADSHEET.TARGET_SHEET_NAME}' 탭이 없습니다.`, logs };
    }

    sheet.clear();
    const sheetData = [
      [`창고코드: ${targetWh}`, `동기화 시간: ${nowTimestamp}`, `전품목 덤프`, ""],
      ["품목코드", "품목명", "규격", "재고수량"]
    ];
    rawItems.forEach(item => {
      sheetData.push([
        item.PROD_CD  || "",
        item.PROD_DES || "",
        item.PROD_SIZE_DES || "",
        Math.round(Number(item.BAL_QTY) || 0)
      ]);
    });
    sheet.getRange(1, 1, sheetData.length, 4).setValues(sheetData);

    return { success:true, message:`동기화 완료 · ${rawItems.length.toLocaleString()}건`, syncTime:nowTimestamp, count:rawItems.length, logs };
  } catch(error) {
    apiLog.responseRaw = { error: error.toString() };
    logs.push(apiLog);
    return { success:false, message:"데이터 처리 중 오류 발생", logs };
  }
}

// ══════════════════════════════════════════════════════════════════
//  규격명 정제: PROD_DES 제거 + 대괄호 제거
// ══════════════════════════════════════════════════════════════════
function cleanSpec_(prodSizeDes, prodDes) {
  let s = (prodSizeDes || "").toString().trim();
  if (prodDes) s = s.replace(prodDes, "").trim();
  s = s.replace(/^[\[\(]+/, "").replace(/[\]\)]+$/, "").trim();
  return s;
}

// ══════════════════════════════════════════════════════════════════
//  [조회 엔진] 시트 캐시 → 전방매칭 → 결과 반환
//  ※ 사이즈/컬러 레이블 추출은 프론트에서 처리 (공통접두사 제거 방식)
// ══════════════════════════════════════════════════════════════════
function fetchInventoryFromSheetCache(whCd, prefix, prodCdInput) {
  whCd   = (whCd   || "").trim();
  prefix = (prefix || "").trim();
  const searchTerms = (prodCdInput || "").split(/\s+/).filter(t => t.length > 0);

  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET.ID);
    const sheet = ss.getSheetByName(CONFIG.SPREADSHEET.TARGET_SHEET_NAME);
    if (!sheet) return { success:false, message:"캐시 데이터 탭이 없습니다." };

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return { success:false, message:"캐시된 재고 정보가 없습니다. 동기화를 먼저 실행하세요." };

    const raw = sheet.getRange(3, 1, lastRow-2, 4).getValues();
    const finalResults = [];
    const logs = [];

    searchTerms.forEach(term => {
      let queryTarget = term;
      if (prefix && !queryTarget.startsWith(prefix)) queryTarget = prefix + queryTarget;

      const matched = [];
      for (let j=0; j<raw.length; j++) {
        const prodCd  = raw[j][0].toString().trim();
        const prodDes = raw[j][1].toString().trim();
        const sizeDes = raw[j][2].toString().trim();
        const qty     = Math.round(Number(raw[j][3]) || 0);

        if (!prodCd.startsWith(queryTarget)) continue;

        // 품번 분해: PREFIX(0)-MODEL(1)-GEN(2)-SIZECODE(3)-COLORCODE(4)
        const parts     = prodCd.split("-");
        const sizeCode  = parts[3] || "";
        const colorCode = parts[4] || "";
        const spec      = cleanSpec_(sizeDes, prodDes);

        matched.push({ prodCd, prodDes, sizeCode, colorCode, spec, qty });
      }

      finalResults.push({ keyword: term, items: matched });
      logs.push({ msg: `매칭 완료 · 창고 ${whCd} · PREFIX ${prefix} · 타겟 ${queryTarget} · ${matched.length}건` });
    });

    return {
      success:true,
      data: finalResults,
      localLogs: logs,
      currentSyncTime: getLastSyncTimeInfo().timeStr
    };
  } catch(e) {
    return { success:false, message:"검색 처리 중 오류: " + e.toString() };
  }
}
