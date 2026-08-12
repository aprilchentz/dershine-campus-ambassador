/**
 * 校園大使報名表單 → Google Sheet（履歷檔案 → Google Drive）
 *
 * 部署（spec §10.1）：
 *   1. 開試算表「校園大使報名_2026」→ 擴充功能 → Apps Script → 貼上這整份
 *   2. 開一個 Google Drive 資料夾放履歷（例如「校園大使_履歷」），
 *      複製資料夾網址裡的 ID → 貼進下面 DRIVE_FOLDER_ID
 *   3. 部署 → 新增部署作業 → 網頁應用程式
 *      執行身分：我 ／ 具有存取權的使用者：任何人
 *   4. 複製 /exec 網址 → 貼進 index.html 的 ENDPOINT 常數
 *   5. 用瀏覽器直接開那個網址，看到 {"ok":true,"msg":"alive"} 才算成功
 *
 * Sheet 第一列標題已建好（順序可調，字串一個字都不能差）：
 * 姓名 Email 生日 歐洲就讀學校 歐洲科系 歐洲年級學期 國家與城市 歐洲就讀身份 畢業或交換結束時間
 * 台灣就讀學校 台灣科系 台灣就讀狀態 台灣就讀狀態補充 任務期間是否全程在歐洲
 * IG帳號 IG追蹤數 IG可否設為公開 Threads帳號 FB破萬社團 FB社團名稱 Dcard帳號 其他平台
 * 作品連結 相關經驗 為什麼想成為校園大使 每週投入時間 月會可配合時段 現有合作關係
 * 追蹤的留歐KOL 得知管道 其他補充 履歷連結 同意條款 送出時間
 *
 * 「履歷連結」是新欄位 —— 表單裡的履歷上傳為選填，沒上傳時這欄會是空字串。
 * 「就讀資訊」欄位 2026/08 改版拆成歐洲／台灣兩組；若舊試算表還留著「就讀學校」
 * 「科系」「年級」三欄，這三欄會停止收到資料，需另外建上面列出的新欄位。
 * 「所在國家與城市」欄位已改名為「國家與城市」並併入歐洲就讀資訊組 —— 若舊試算表
 * 還留著「所在國家與城市」，需把該欄標題改成「國家與城市」，否則這欄會停止收到資料。
 */

const SHEET_ID   = '1vp86LlglX0k91x6w42rgUNKke0QnUBKgoxS8PLCU4NU';
const SHEET_NAME = '';   // 留空 = 用第一個工作表。CSV 匯入產生的分頁不叫「工作表1」，
                         // 寫死名字是這個串接最常見的失敗點，所以預設不寫死。

const DRIVE_FOLDER_ID = '';   // ⚠️ 部署前必填：放履歷檔案的 Drive 資料夾 ID
const RESUME_MAX_BYTES = 10 * 1024 * 1024;   // 與 index.html 的 RESUME_MAX_BYTES 對齊

function _sheet(){
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return (SHEET_NAME && ss.getSheetByName(SHEET_NAME)) || ss.getSheets()[0];
}

// 履歷 base64 → 存進 Drive 資料夾 → 回傳檔案連結。沒有 resume 或沒設定資料夾就回傳空字串。
function _saveResume(data){
  if (!data.resume || !data.resume.base64) return '';
  if (!DRIVE_FOLDER_ID) throw new Error('DRIVE_FOLDER_ID 尚未設定，無法儲存履歷');

  const bytes = Utilities.base64Decode(data.resume.base64);
  if (bytes.length > RESUME_MAX_BYTES) throw new Error('履歷檔案超過大小上限');

  const safeName = String(data.resume.name || 'resume').replace(/[\/\\]/g, '_');
  const fileName = `${data['姓名'] || '未填姓名'}_${new Date().getTime()}_${safeName}`;
  const blob = Utilities.newBlob(bytes, data.resume.mime || 'application/octet-stream', fileName);
  const file = DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(blob);
  return file.getUrl();
}

function doPost(e){
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(20000);                       // 防止同時送出時覆蓋同一列
    const data = JSON.parse(e.postData.contents);
    if (data._hp) return _json({ok:true});      // 蜜罐命中，假成功不寫入

    data['履歷連結'] = _saveResume(data);

    const sh = _sheet();
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const row = headers.map(h => {
      if (h === '送出時間') return new Date();
      const v = data[h];
      return Array.isArray(v) ? v.join('、') : (v == null ? '' : v);   // 多選轉字串
    });
    sh.appendRow(row);
    return _json({ok:true});
  }catch(err){
    return _json({ok:false, error:String(err)});
  }finally{
    lock.releaseLock();
  }
}

function doGet(){ return _json({ok:true, msg:'alive'}); }   // 用來測試部署是否成功

function _json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 部署前先在編輯器裡跑這個（選 testWrite → 執行），會寫入一列測試資料。
 * 通過代表 SHEET_ID 與權限都對，之後把那一列刪掉即可。
 */
function testWrite(){
  const sh = _sheet();
  Logger.log('工作表名稱：' + sh.getName() + '／欄數：' + sh.getLastColumn());
  doPost({ postData: { contents: JSON.stringify({ 姓名:'測試', Email:'test@example.com' }) } });
}
