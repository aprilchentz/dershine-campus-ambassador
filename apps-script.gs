/**
 * 校園大使報名表單 → Google Sheet（履歷檔案 → Google Drive）
 *
 * ⚠️ 整套（試算表／Apps Script／履歷資料夾）都用 Der Shine 顧問 Gmail 執行，
 *    不要掛在個人帳號下。原因：確認信的寄件人就是執行腳本的帳號，
 *    用機構帳號就不必去設 Gmail 別名（別名要走 outlook SMTP，容易被判垃圾信）；
 *    履歷也存在機構的 Drive，之後交接不用重弄。
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
 * 【收不到資料時】先跑 checkSetup（不寫入任何東西，只檢查設定），
 * 它會一次列出分頁、標題列、Drive 三處的問題。
 * 改完程式碼一定要重新部署：Deploy → Manage deployments → 鉛筆 → New version，
 * 只是存檔的話 /exec 跑的還是舊版。
 *
 * 【看到 "Unexpected error while getting the method or property ... on object DriveApp"】
 * 這不是資料夾 ID 的問題，是 Drive 權限沒授權到 —— 專案先前只授權了試算表權限，
 * 後來才加入 DriveApp，重跑時不一定會自動要求補授權。解法：
 *   1. 專案設定（齒輪）→ 勾選「顯示 appsscript.json 資訊清單檔案」
 *   2. 打開 appsscript.json，補上本專案的 oauthScopes（同資料夾的 appsscript.json 可直接參考）：
 *        "oauthScopes": [
 *          "https://www.googleapis.com/auth/spreadsheets",
 *          "https://www.googleapis.com/auth/drive",
 *          "https://www.googleapis.com/auth/script.send_mail",
 *          "https://www.googleapis.com/auth/gmail.send",
 *          "https://www.googleapis.com/auth/gmail.settings.basic"
 *        ]
 *   3. 存檔 → 重跑 checkSetup → 會跳出新的同意畫面，全部允許
 *   4. 重新部署（New version）
 *
 * 注意：一旦在 appsscript.json 明列 oauthScopes，Apps Script 就不再自動偵測權限，
 * 少列的權限一律報錯。本檔案用到的服務與對應權限：
 *   SpreadsheetApp → spreadsheets ／ DriveApp → drive
 *   MailApp → script.send_mail ／ GmailApp.sendEmail → gmail.send
 *   GmailApp.getAliases → gmail.settings.basic
 * 之後若新增其他服務，記得一併補進 oauthScopes。
 *
 * Sheet 第一列標題已建好（順序可調，字串一個字都不能差）：
 * 姓名 Email 生日 歐洲就讀學校 歐洲科系 歐洲年級學期 國家與城市 歐洲就讀身份 畢業或交換結束時間
 * 台灣就讀學校 台灣科系 台灣就讀狀態 台灣就讀狀態補充 任務期間是否全程在歐洲
 * IG帳號 IG追蹤數 IG可否設為公開 Threads帳號 FB破萬社團 FB社團名稱 Dcard帳號 其他平台
 * 作品連結 相關經驗 為什麼想成為校園大使 每週投入時間 月會可配合時段 現有合作關係 合作關係說明
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

// ── 自動確認信 ────────────────────────────────
// 設 false 就完全不寄。開啟後需重新授權（多了寄信權限）並重新部署。
// 免費 Gmail 帳號每天 100 封上限；用 Workspace 帳號是 1500 封。
const SEND_CONFIRMATION = true;
const FROM_NAME    = 'Der Shine 留德青山';
const NOTIFY_EMAIL = '';   // 選填：每有一筆報名就通知你自己，留空不寄

// 寄件地址。留空 = 用執行腳本的帳號（你的 Gmail）。
// 要填別的地址（例如 der-shine@outlook.com），必須先在 Gmail 設定成驗證過的別名：
//   Gmail 設定 → 查看所有設定 → 帳戶和匯入 → 「以這個地址寄送郵件」→ 新增
//   → 收 outlook 信箱裡的驗證碼 → 輸入完成
// 沒設成別名就填在這裡的話，Gmail 會直接拒絕，信一封都寄不出去。
// 設定完可以跑 checkSetup 確認別名有沒有被認出來。
const FROM_ALIAS = '';   // 例：'der-shine@outlook.com'

// 申請人按「回覆」時會寄到哪。留空就跟著寄件地址走。
const REPLY_TO = '';

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

    // 履歷存檔失敗不能連帶讓整份報名消失 —— 報名內容照寫，
    // 「履歷連結」改記下失敗原因，之後從那一列就看得出要回頭跟誰要檔案。
    let resumeErr = '';
    try{
      data['履歷連結'] = _saveResume(data);
    }catch(err){
      resumeErr = String(err && err.message || err);
      data['履歷連結'] = '⚠️ 履歷未存檔：' + resumeErr;
    }

    const sh = _sheet();
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const row = headers.map(h => {
      if (h === '送出時間') return new Date();
      const v = data[h];
      return Array.isArray(v) ? v.join('、') : (v == null ? '' : v);   // 多選轉字串
    });
    sh.appendRow(row);

    // 寄信放在寫入之後，且獨立 try —— 信寄不出去（配額用完、信箱打錯）
    // 絕對不能讓已經寫進試算表的報名回報成失敗，害對方重送一次
    try{ _sendMails(data); }catch(err){ Logger.log('寄信失敗：' + err); }

    // 回報 ok:true（報名已存下），但把履歷的問題一併帶回去，方便排查
    return _json(resumeErr ? {ok:true, warn:'履歷未存檔：' + resumeErr} : {ok:true});
  }catch(err){
    return _json({ok:false, error:String(err)});
  }finally{
    lock.releaseLock();
  }
}

function doGet(){ return _json({ok:true, msg:'alive'}); }   // 用來測試部署是否成功

/**
 * 統一的寄信出口。
 * 有設 FROM_ALIAS 就走 GmailApp —— 只有它支援指定寄件地址，
 * 而且只吃 Gmail 裡驗證過的別名。沒設就退回 MailApp（寄件人＝執行帳號）。
 */
function _send(to, subject, htmlBody){
  if (FROM_ALIAS){
    GmailApp.sendEmail(to, subject, '', {
      htmlBody: htmlBody,
      name: FROM_NAME,
      from: FROM_ALIAS,
      replyTo: REPLY_TO || FROM_ALIAS
    });
  } else {
    MailApp.sendEmail({
      to: to, subject: subject, htmlBody: htmlBody,
      name: FROM_NAME,
      replyTo: REPLY_TO || undefined
    });
  }
}

// 申請人確認信 ＋（選填）給自己的通知信
function _sendMails(data){
  const to   = String(data['Email'] || '').trim();
  const name = String(data['姓名'] || '').trim() || '你';

  if (SEND_CONFIRMATION && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)){
    _send(to, '我們收到你的報名了 ｜ Der Shine 青山校園大使 2026',
        '<div style="font-family:-apple-system,\'Noto Sans TC\',sans-serif;font-size:15px;line-height:1.9;color:#22372b;max-width:520px">'
      + '<p>' + _esc(name) + ' 你好，</p>'
      + '<p>我們已經收到你的<b>青山校園大使</b>報名表了，謝謝你願意把時間花在這件事上。</p>'
      + '<p>接下來我們會逐一看過每一份報名，<b>先到先審</b>。'
      + '報名在 <b>9/15</b> 截止，我們會在審閱後主動與你聯繫，不需要另外來信詢問進度。</p>'
      + '<p>如果你臨時想補充作品或更新資料，直接回覆這封信就可以。</p>'
      + '<p style="margin-top:28px">Der Shine 留德青山</p>'
      + '<p style="font-size:12px;color:#6b7d70">這是系統自動寄出的確認信，代表我們收到了你的表單。</p>'
      + '</div>');
  }

  if (NOTIFY_EMAIL){
    _send(NOTIFY_EMAIL, '新報名：' + name + '（' + (data['歐洲就讀學校'] || '未填學校') + '）',
        '<p>IG：' + _esc(data['IG帳號'] || '—')
      + '　追蹤數：' + _esc(data['IG追蹤數'] || '—') + '</p>'
      + '<p>履歷：' + _esc(data['履歷連結'] || '未上傳') + '</p>'
      + '<p><a href="https://docs.google.com/spreadsheets/d/' + SHEET_ID + '">開啟試算表</a></p>');
  }
}

function _esc(s){
  return String(s).replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
}

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
  // Email 留空：確認信只寄給格式正確的信箱，測試時不會真的寄信出去、也不吃寄信配額
  doPost({ postData: { contents: JSON.stringify({ 姓名:'測試', Email:'' }) } });
}

/**
 * 表單沒收到資料時，先跑這個（選 checkSetup → 執行 → 看執行紀錄）。
 * 不寫入任何資料，只檢查設定，一次把所有問題列出來。
 */
function checkSetup(){
  const out = [];
  // 結果統一在 finally 印出 —— 任何一段意外中斷，前面已經查到的結果都還看得到。
  // （只在最後 Logger.log 的話，第 4 段一 throw 就連 1–3 段的結果都一起消失）
  try{ _checkSetup(out); }
  catch(err){ out.push('❌ 檢查中斷：' + err); }
  finally{ Logger.log(out.join('\n')); }
}

function _checkSetup(out){
  // 1. 試算表與分頁 —— SHEET_NAME 留空時用第一個分頁，
  //    報名表頭不在第一個分頁時，資料會安靜地寫到隔壁分頁去
  let sh;
  try{
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const names = ss.getSheets().map(s => s.getName());
    sh = _sheet();
    out.push('✅ 試算表：' + ss.getName());
    out.push('   全部分頁：' + names.join('｜'));
    out.push((names.length > 1 ? '⚠️ ' : '✅ ') + '實際寫入分頁：' + sh.getName()
             + (names.length > 1 ? '（有多個分頁，請確認報名表頭就在這一個）' : ''));
  }catch(err){
    out.push('❌ 打不開試算表：' + err);
    out.push('   目前的 SHEET_ID：' + SHEET_ID + '（' + SHEET_ID.length + ' 字元）');
    // 原生 Google 試算表的 ID 是 44 字元；33 字元通常是「上傳的檔案」，
    // 例如直接丟上 Drive 但沒轉成 Google 試算表格式的 .xlsx —— SpreadsheetApp 打不開這種檔案。
    if (SHEET_ID.length !== 44){
      out.push('   ⚠️ 原生 Google 試算表的 ID 是 44 字元。長度不對的話，多半是');
      out.push('      （a）貼到了別的檔案的 ID，或（b）那是上傳的 .xlsx，還沒轉成 Google 試算表格式');
      out.push('      （在 Drive 上按右鍵 → 開啟工具 → Google 試算表 → 檔案 → 儲存為 Google 試算表）');
    }
    out.push('   正確的 ID 取法：開啟試算表，複製網址中 /d/ 與 /edit 之間那一段。');
    return;   // 開不了試算表，後面的標題列檢查沒有意義
  }

  // 2. 標題列 —— 欄位靠字串比對，差一個字該欄就永遠是空的
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const REQUIRED = ['姓名','Email','生日','歐洲就讀學校','歐洲科系','歐洲年級學期','國家與城市',
    '歐洲就讀身份','畢業或交換結束時間','台灣就讀學校','台灣科系','台灣就讀狀態','台灣就讀狀態補充',
    '任務期間是否全程在歐洲','IG帳號','IG追蹤數','IG可否設為公開','Threads帳號','FB破萬社團',
    'FB社團名稱','Dcard帳號','其他平台','作品連結','相關經驗','為什麼想成為校園大使','每週投入時間',
    '月會可配合時段','現有合作關係','合作關係說明','追蹤的留歐KOL','得知管道','其他補充','履歷連結',
    '同意條款','送出時間'];
  const missing = REQUIRED.filter(h => headers.indexOf(h) === -1);
  const extra   = headers.filter(h => h && REQUIRED.indexOf(h) === -1);
  out.push(missing.length ? '❌ 缺少欄位標題（這些欄不會收到資料）：' + missing.join('、')
                          : '✅ 標題列 34 欄齊全');
  if (extra.length) out.push('ℹ️ 表單不會寫入的欄位（舊欄位可刪）：' + extra.join('、'));

  // 3. Drive —— 分兩段測，因為兩種失敗的解法完全不同：
  //    先用不需要 ID 的 getRootFolder() 確認「服務本身能不能用」，
  //    通過了再測 ID。不分開測的話，授權問題會被誤讀成 ID 打錯。
  let driveOk = false;
  try{
    DriveApp.getRootFolder().getName();
    driveOk = true;
    out.push('✅ DriveApp 服務可用（Drive 權限已授權）');
  }catch(err){
    out.push('❌ DriveApp 服務載入失敗 —— 這與資料夾 ID 無關，是 Drive 權限沒授權到。');
    out.push('   訊息：' + err);
    out.push('   解法：在 appsscript.json 明確宣告 oauthScopes（見檔案開頭說明），');
    out.push('        存檔後重跑本函式，會跳出新的同意畫面 → 全部允許。');
  }

  if (!DRIVE_FOLDER_ID){
    out.push('❌ DRIVE_FOLDER_ID 沒填，履歷一律存不進去');
  } else if (driveOk){
    try{
      const f = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      out.push('✅ Drive 資料夾：' + f.getName() + '（ID 正確）');
    }catch(err){
      out.push('❌ 資料夾 ID 有問題：' + err);
      out.push('   跑 listMyFolders() 可以列出你的資料夾與正確 ID。');
    }
  }

  // 4. 寄件地址 —— 別名沒驗證過的話，Gmail 會拒收整封信，
  //    確認信會一封都寄不出去，所以上線前一定要在這裡看到 ✅
  if (!SEND_CONFIRMATION){
    out.push('ℹ️ SEND_CONFIRMATION = false，不會寄確認信');
  } else if (!FROM_ALIAS){
    // 讀執行帳號的信箱要 userinfo.email 權限。那是純資訊、不值得為它多要一個權限，
    // 所以讀不到就算了 —— 絕不能讓一行說明文字害整個檢查中斷。
    let who = '';
    try{ who = '（' + Session.getEffectiveUser().getEmail() + '）'; }catch(err){}
    out.push('ℹ️ 寄件地址：執行腳本的帳號' + who);
  } else {
    try{
      const aliases = GmailApp.getAliases();
      if (aliases.indexOf(FROM_ALIAS) !== -1){
        out.push('✅ 寄件地址：' + FROM_ALIAS + '（別名已驗證）');
      } else {
        out.push('❌ ' + FROM_ALIAS + ' 不在已驗證的別名清單裡，信會寄不出去。');
        out.push('   目前可用的別名：' + (aliases.length ? aliases.join('、') : '（沒有）'));
        out.push('   請到 Gmail 設定 → 帳戶和匯入 → 「以這個地址寄送郵件」新增並完成驗證。');
      }
    }catch(err){
      out.push('❌ 讀不到 Gmail 別名清單（gmail.settings.basic 權限沒授權）：' + err);
    }
  }
}

/**
 * 找不到正確的資料夾 ID 時跑這個 —— 直接列出你 Drive 裡最近的資料夾與 ID，
 * 不用從網址上剪字串（網址上常常還黏著 ?usp=sharing 或 /u/0/ 這類東西）。
 */
function listMyFolders(){
  const out = [];
  const it = DriveApp.getFolders();
  let n = 0;
  while (it.hasNext() && n < 40){
    const f = it.next();
    out.push(f.getName() + '\n    ' + f.getId());
    n++;
  }
  Logger.log(out.length ? out.join('\n') : '（找不到任何資料夾）');
}
