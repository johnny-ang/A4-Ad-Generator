/**
 * 電梯廣告編輯工具 — Apps Script 代理層
 * ------------------------------------------------------------
 * 用途：
 *   1. queryStaff  — 依關鍵字查詢 Google Sheet 人員資料（前台身分驗證）
 *   2. getConfig   — 讀取 GitHub 上目前的 config.json（後台載入現況）
 *   3. updateConfig— 後台編輯完成後，代理寫回 GitHub（前端不接觸 Token）
 *
 * 部署方式：
 *   Google Apps Script 專案 → 部署 → 新增部署作業 → 網頁應用程式
 *   執行身分：我　　存取權限：任何人
 *   部署後取得的網址即為前台/後台呼叫的 API base URL
 *
 * 必要設定（App Script 編輯器 → 專案設定 →指令碼屬性）：
 *   GITHUB_TOKEN   — GitHub Personal Access Token（僅需 repo 內容讀寫權限）
 *   GITHUB_OWNER   — GitHub 使用者/組織名稱
 *   GITHUB_REPO    — repo 名稱
 *   GITHUB_BRANCH  — 分支名稱，預設 main
 *   SHEET_ID       — 人員資料 Google Sheet 的試算表 ID
 *   ADMIN_EMAILS   — 允許寫入後台設定的 Google 帳號白名單，以逗號分隔（選填但建議設定）
 */

const CONFIG_PATH_IN_REPO = 'config.json'; // GitHub repo 內 config 檔案路徑
const STAFF_SHEET_NAME = '人員資料';        // Google Sheet 分頁名稱

/* ============================================================
   入口：GET（查詢類）
   ============================================================ */
function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'queryStaff') {
      return respond(queryStaff(e.parameter.keyword || ''));
    }
    if (action === 'getConfig') {
      return respond(getConfigFromGithub());
    }
    return respond({ error: '未知的 action：' + action }, 400);
  } catch (err) {
    return respond({ error: err.message }, 500);
  }
}

/* ============================================================
   入口：POST（寫入類，目前僅 updateConfig）
   ============================================================ */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action !== 'updateConfig') {
      return respond({ error: '未知的 action：' + body.action }, 400);
    }

    // 簡易白名單檢查（選填）：body 需帶入後台登入者的 email
    const whitelist = getScriptProp_('ADMIN_EMAILS');
    if (whitelist) {
      const allowed = whitelist.split(',').map(s => s.trim());
      if (!body.editorEmail || allowed.indexOf(body.editorEmail) === -1) {
        return respond({ error: '此帳號無後台編輯權限' }, 403);
      }
    }

    const result = updateConfigOnGithub(body.config);
    return respond({ success: true, commitSha: result.commit.sha });
  } catch (err) {
    return respond({ error: err.message }, 500);
  }
}

/* ============================================================
   人員資料查詢（對應前台關鍵字驗證）
   ============================================================ */
function queryStaff(keyword) {
  if (!keyword) return { found: false };
  const sheetId = getScriptProp_('SHEET_ID');
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(STAFF_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  // 首欄（A欄）＝查詢關鍵字，第一列為標題列
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]).trim() === keyword.trim()) {
      return {
        found: true,
        staff: {
          keyword: row[0],
          name: row[1],
          phone: row[2],
          brand: row[3],
          branch: row[4],
          photo: row[5],
          qrcode: row[6] || ''
        }
      };
    }
  }
  return { found: false };
}

/* ============================================================
   GitHub Contents API 存取
   ============================================================ */
function githubApiBase_() {
  const owner = getScriptProp_('GITHUB_OWNER');
  const repo = getScriptProp_('GITHUB_REPO');
  return `https://api.github.com/repos/${owner}/${repo}/contents/`;
}

function githubHeaders_() {
  return {
    Authorization: 'Bearer ' + getScriptProp_('GITHUB_TOKEN'),
    Accept: 'application/vnd.github+json'
  };
}

function getConfigFromGithub() {
  const branch = getScriptProp_('GITHUB_BRANCH') || 'main';
  const url = githubApiBase_() + CONFIG_PATH_IN_REPO + '?ref=' + branch;
  const res = UrlFetchApp.fetch(url, { headers: githubHeaders_(), muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('讀取 GitHub 上的 config.json 失敗：' + res.getContentText());
  }
  const fileInfo = JSON.parse(res.getContentText());
  const content = Utilities.newBlob(Utilities.base64Decode(fileInfo.content)).getDataAsString();
  return { config: JSON.parse(content), sha: fileInfo.sha };
}

function updateConfigOnGithub(newConfig) {
  const branch = getScriptProp_('GITHUB_BRANCH') || 'main';
  // GitHub 更新檔案必須先取得目前檔案的 sha
  const current = getConfigFromGithub();
  const url = githubApiBase_() + CONFIG_PATH_IN_REPO;
  const payload = {
    message: '後台更新版型設定 ' + new Date().toISOString(),
    content: Utilities.base64Encode(JSON.stringify(newConfig, null, 2)),
    sha: current.sha,
    branch: branch
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'put',
    headers: githubHeaders_(),
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    throw new Error('寫入 GitHub 失敗：' + res.getContentText());
  }
  return JSON.parse(res.getContentText());
}

/* ============================================================
   工具函式
   ============================================================ */
function getScriptProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function respond(obj, code) {
  // Apps Script Web App 無法自訂 HTTP status code，錯誤仍以 200 回傳，
  // 前端請自行檢查回傳內容中是否有 error 欄位
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
