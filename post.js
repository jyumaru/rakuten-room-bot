const { Builder, By, until, Key } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const { google } = require('googleapis');
const fs = require('fs');

// ============================================================
// 設定
// ============================================================
const CONFIG = {
  EMAIL: process.env.EM,
  PASS:  process.env.PW,
  SHEET: process.env.SID,
  SHEET_TAB: '楽天ROOM投稿リスト',
  MAX_ITEMS_PER_RUN: 5,
  WAIT_BETWEEN_POSTS_MS: 5000,
};

const COL = {
  CREATED_AT: 0, CATEGORY: 1, ITEM_CODE: 2, ITEM_NAME: 3,
  PRICE: 4, ITEM_URL: 5, POST_TEXT: 6, STATUS: 7,
};

const BANNED_KEYWORDS = [
  '腱鞘炎', '肩こり', '腰痛', '頭痛', '眼精疲労', '疲労回復', '不眠',
  '冷え性', 'むくみ', 'ダイエット効果', 'うつ',
  '治る', '治療', '改善', '予防', '効果', '効能', '症状',
  '医療', '医薬', '診断', '処方',
  '最安', '最安値', '激安', '送料無料確約',
];

function sanitizeContent(text) {
  let cleaned = text || '';
  const hits = [];
  for (const word of BANNED_KEYWORDS) {
    if (cleaned.includes(word)) {
      hits.push(word);
      cleaned = cleaned.split(word).join('');
    }
  }
  cleaned = cleaned.replace(/[ 　]{2,}/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
  return { cleaned, hits };
}

function extractBannedFromApiMessage(message) {
  if (!message) return [];
  const m = message.match(/禁止されたキーワードが含まれます[:：]\s*(.+?)(?:$|。|\s)/);
  if (!m) return [];
  return m[1].split(/[、,・\s]+/).filter(Boolean);
}

function buildRoomUrl(itemCode) {
  if (!itemCode || !itemCode.includes(':')) return null;
  return `https://room.rakuten.co.jp/mix?itemcode=${encodeURIComponent(itemCode)}&scid=we_room_upc60`;
}

// ============================================================
// Googleスプレッドシート
// ============================================================
async function getSheetsClient() {
  const b64 = process.env.GSA;
  const json = Buffer.from(b64, 'base64').toString('utf8');
  const credentials = JSON.parse(json);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

async function getUnpostedItems() {
  try {
    const sheets = await getSheetsClient();
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.SHEET,
      range: `${CONFIG.SHEET_TAB}!A:H`,
    });
    const rows = res.data.values || [];
    const items = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const status = (row[COL.STATUS] || '').toString().trim();
      if (status === '済' || status === 'スキップ') continue;
      items.push({
        rowIndex:  i + 1,
        category:  row[COL.CATEGORY]  || '',
        itemCode:  row[COL.ITEM_CODE] || '',
        itemName:  row[COL.ITEM_NAME] || '',
        itemUrl:   row[COL.ITEM_URL]  || '',
        postText:  row[COL.POST_TEXT] || '',
        status:    status,
      });
    }
    return items;
  } catch (e) {
    console.error('スプレッドシート取得エラー:', e.message);
    return [];
  }
}

async function updateStatus(rowIndex, statusText) {
  try {
    const sheets = await getSheetsClient();
    await sheets.spreadsheets.values.update({
      spreadsheetId: CONFIG.SHEET,
      range: `${CONFIG.SHEET_TAB}!H${rowIndex}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[statusText]] },
    });
    console.log(`行${rowIndex}を「${statusText}」に更新`);
  } catch (e) {
    console.error(`行${rowIndex}の更新エラー:`, e.message);
  }
}

// ============================================================
// Selenium
// ============================================================
function createDriver() {
  const options = new chrome.Options();
  options.addArguments('--headless=new');
  options.addArguments('--no-sandbox');
  options.addArguments('--disable-dev-shm-usage');
  options.addArguments('--disable-gpu');
  options.addArguments('--window-size=1280,900');
  options.addArguments('--lang=ja-JP');
  options.addArguments(
    '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) ' +
    'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
  );
  return new Builder().forBrowser('chrome').setChromeOptions(options).build();
}

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

async function screenshot(driver, filename) {
  try {
    const img = await driver.takeScreenshot();
    fs.writeFileSync(filename, img, 'base64');
  } catch (e) {}
}

// ============================================================
// ログイン
// ============================================================
async function performLogin(driver) {
  console.log('ログイン処理開始...');
  let emailInput = null;
  for (let attempt = 0; attempt < 5 && !emailInput; attempt++) {
    for (const sel of [By.name('u'), By.css('input[type="email"]'), By.css('input[type="text"]')]) {
      try {
        await driver.wait(until.elementLocated(sel), 2000);
        emailInput = await driver.findElement(sel);
        if (emailInput) break;
      } catch (e) {}
    }
    if (!emailInput) await sleep(1500);
  }
  if (!emailInput) { console.log('メール入力欄なし'); return false; }

  await emailInput.clear();
  await emailInput.sendKeys(CONFIG.EMAIL);
  await sleep(800);
  try {
    const btn = await driver.findElement(By.css('button[type="submit"], input[type="submit"]'));
    await driver.executeScript('arguments[0].click();', btn);
  } catch (e) { await emailInput.sendKeys(Key.RETURN); }
  await sleep(3000);

  let passInput = null;
  for (let attempt = 0; attempt < 5 && !passInput; attempt++) {
    try { passInput = await driver.findElement(By.css('input[type="password"]')); }
    catch (e) { await sleep(1500); }
  }
  if (!passInput) { console.log('パスワード入力欄なし'); return false; }

  await passInput.clear();
  await passInput.sendKeys(CONFIG.PASS);
  await sleep(800);
  try {
    const btn = await driver.findElement(By.css('button[type="submit"], input[type="submit"]'));
    await driver.executeScript('arguments[0].click();', btn);
  } catch (e) { await passInput.sendKeys(Key.RETURN); }
  await sleep(5000);

  const url = await driver.getCurrentUrl();
  if (url.includes('rakuten.co.jp')) {
    console.log(`✅ ログイン成功: ${url}`);
    return true;
  }
  console.log(`ログイン後URL（想定外）: ${url}`);
  return false;
}

// ============================================================
// ROOM投稿: XHR + fetch 両方フック、URL変化も判定材料に
// ============================================================
async function postToRoom(driver, roomUrl, postText) {
  console.log(`ROOM投稿ページへ: ${roomUrl}`);
  await driver.get(roomUrl);
  await sleep(4500);

  const beforeUrl = await driver.getCurrentUrl();
  if (beforeUrl.includes('login.account.rakuten.com') || beforeUrl.includes('login.rakuten.co.jp')) {
    return { status: 'login_required' };
  }

  // OKポップアップを閉じる
  for (let i = 0; i < 3; i++) {
    try {
      const okBtn = await driver.findElement(By.xpath('//button[normalize-space(text())="OK"]'));
      await driver.executeScript('arguments[0].click();', okBtn);
      console.log(`OKクリック ${i + 1}回目`);
      await sleep(800);
    } catch (e) { break; }
  }

  // textareaの出現を待つ
  try {
    await driver.wait(until.elementLocated(By.id('collect-content')), 8000);
  } catch (e) {
    await screenshot(driver, `no_form_${Date.now()}.png`);
    return { status: 'form_not_found' };
  }

  const itemCodeMatch = roomUrl.match(/itemcode=([^&]+)/);
  const itemCode = itemCodeMatch ? decodeURIComponent(itemCodeMatch[1]) : '';
  console.log(`itemCode: ${itemCode}`);

  // XHR + fetch 両方フック、collect() 実行、captureを返す
  const apiResultRaw = await driver.executeAsyncScript(`
    const callback = arguments[arguments.length - 1];
    const content = arguments[0];
    const itemCode = arguments[1];
    let captured = null;
    const log = [];

    // fetch フック
    const origFetch = window.fetch;
    window.fetch = async function(url, opts) {
      const res = await origFetch.apply(this, arguments);
      try {
        const urlStr = typeof url === 'string' ? url : (url && url.url ? url.url : '');
        if (urlStr.includes('/api/collect') || urlStr.includes('/collect')) {
          log.push('fetch captured: ' + urlStr);
          const clone = res.clone();
          const text = await clone.text();
          captured = { source: 'fetch', status: res.status, body: text };
        }
      } catch (e) { log.push('fetch hook error: ' + e.message); }
      return res;
    };

    // XHR フック
    const origXhrOpen = XMLHttpRequest.prototype.open;
    const origXhrSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function(method, url) {
      this._collectUrl = String(url);
      return origXhrOpen.apply(this, arguments);
    };
    XMLHttpRequest.prototype.send = function() {
      const xhr = this;
      xhr.addEventListener('loadend', function() {
        try {
          if (xhr._collectUrl && (xhr._collectUrl.includes('/api/collect') || xhr._collectUrl.includes('/collect'))) {
            log.push('xhr captured: ' + xhr._collectUrl + ' status=' + xhr.status);
            captured = { source: 'xhr', status: xhr.status, body: xhr.responseText };
          }
        } catch (e) { log.push('xhr hook error: ' + e.message); }
      });
      return origXhrSend.apply(this, arguments);
    };

    // textareaに入力
    const ta = document.querySelector('#collect-content');
    if (ta) {
      ta.value = content;
      ta.dispatchEvent(new Event('input',  { bubbles: true }));
      ta.dispatchEvent(new Event('change', { bubbles: true }));
    }

    // AngularJS scope から collect() 呼び出し
    let invoked = false;
    try {
      if (typeof angular !== 'undefined') {
        const el = document.querySelector('[ng-controller], [data-ng-controller]') || document.body;
        const scope = angular.element(el).scope();
        if (scope && typeof scope.collect === 'function') {
          if (itemCode && !scope.itemCode) scope.itemCode = itemCode;
          if (scope.collectContent !== undefined) scope.collectContent = content;
          scope.$apply(function() { scope.collect(); });
          invoked = true;
          log.push('collect() invoked');
        } else {
          log.push('no scope or no collect fn');
        }
      } else {
        log.push('no angular');
      }
    } catch (e) {
      log.push('scope error: ' + e.message);
    }

    // collect() 呼び出せなかった時は送信ボタンを探してクリック
    if (!invoked) {
      const candidates = Array.from(document.querySelectorAll('button, input[type="submit"], [ng-click*="collect"]'));
      const submitBtn = candidates.find(b =>
        /投稿|collect|送信|post/i.test((b.textContent || '') + ' ' + (b.getAttribute('ng-click') || '') + ' ' + (b.value || ''))
      );
      if (submitBtn) {
        log.push('clicking submit: ' + (submitBtn.textContent || submitBtn.value));
        submitBtn.click();
      } else {
        log.push('no submit button found');
      }
    }

    // 捕捉 or タイムアウトまで待つ
    const start = Date.now();
    const timer = setInterval(function() {
      if (captured || Date.now() - start > 18000) {
        clearInterval(timer);
        callback({
          captured: captured,
          log: log,
          finalUrl: location.href,
          elapsed: Date.now() - start,
        });
      }
    }, 300);
  `, postText, itemCode);

  console.log('内部ログ:');
  if (apiResultRaw && apiResultRaw.log) {
    apiResultRaw.log.forEach(l => console.log(`  | ${l}`));
  }

  return {
    status: 'posted',
    captured: apiResultRaw ? apiResultRaw.captured : null,
    finalUrl: apiResultRaw ? apiResultRaw.finalUrl : '',
    elapsed: apiResultRaw ? apiResultRaw.elapsed : 0,
  };
}

// ============================================================
// 成功判定: (1) API応答 (2) URL変化 のどちらかで成功扱い
// ============================================================
function judgePostResult(postRes) {
  // 明示的にAPIレスポンスを捕捉している場合
  if (postRes.captured) {
    const http = postRes.captured.status;
    let body = {};
    try { body = JSON.parse(postRes.captured.body || '{}'); } catch (e) {}
    const message = body.message || '';

    // 2xx かつ status != 'error' → 正常
    if (http >= 200 && http < 300 && body.status !== 'error') {
      return { ok: true, reason: `api_ok (${postRes.captured.source})`, detail: body };
    }
    // 「重複操作です」= 既に投稿済み → 成功と同等扱い
    if (http === 400 && /重複操作/.test(message)) {
      return { ok: true, reason: 'already_posted', httpStatus: http, message };
    }
    return { ok: false, reason: 'api_error', httpStatus: http, message, body };
  }

  // API応答を捕捉できなかった場合、URL変化で判定
  const finalUrl = postRes.finalUrl || '';
  if (finalUrl.includes('/mix/collect?') || finalUrl.includes('/collect?itemcode=')) {
    return { ok: true, reason: 'url_changed_to_collect', finalUrl };
  }

  return { ok: false, reason: 'no_evidence', finalUrl };
}

// ============================================================
// 1商品を投稿するメインフロー
// ============================================================
async function postItem(driver, item) {
  const roomUrl = buildRoomUrl(item.itemCode);
  if (!roomUrl) {
    console.log(`❌ itemCodeが不正: "${item.itemCode}"`);
    return { status: 'invalid_item_code' };
  }

  let { cleaned, hits } = sanitizeContent(item.postText);
  if (hits.length) console.log(`⚠️ 事前サニタイズ: ${hits.join(', ')} を除去`);
  if (!cleaned) { console.log('❌ 紹介文が空'); return { status: 'empty_post_text' }; }

  console.log('楽天にログイン中...');
  await driver.get('https://my.bookmark.rakuten.co.jp/');
  await sleep(2000);
  const loggedIn = await performLogin(driver);
  if (!loggedIn) return { status: 'login_failed' };

  for (let attempt = 1; attempt <= 2; attempt++) {
    console.log(`紹介文を入力中... (試行 ${attempt}/2)`);
    const postRes = await postToRoom(driver, roomUrl, cleaned);

    if (postRes.status === 'login_required') return { status: 'login_failed' };
    if (postRes.status === 'form_not_found') return { status: 'form_not_found' };

    console.log(`elapsed=${postRes.elapsed}ms finalUrl=${(postRes.finalUrl || '').substring(0, 120)}`);

    const j = judgePostResult(postRes);
    if (j.ok) {
      console.log(`✅ 投稿成功 (${j.reason}): ${item.itemName.substring(0, 40)}`);
      return { status: 'success', reason: j.reason };
    }

    console.log(`❌ 判定: ${j.reason}`);
    if (j.httpStatus) console.log(`   http=${j.httpStatus} message=${j.message}`);

    // 禁止語エラーならリトライ
    if (j.httpStatus === 400 && /禁止されたキーワード/.test(j.message || '')) {
      const extra = extractBannedFromApiMessage(j.message);
      if (extra.length) {
        console.log(`⚠️ API指摘の禁止語を除去: ${extra.join(', ')}`);
        for (const w of extra) cleaned = cleaned.split(w).join('');
        cleaned = cleaned.replace(/[ 　]{2,}/g, ' ').trim();
        continue;
      }
    }

    // no_evidence の場合も、現実的には投稿されている可能性があるので manual_check として扱う
    if (j.reason === 'no_evidence') {
      return { status: 'manual_check', detail: j };
    }
    return { status: 'api_failed', detail: j };
  }
  return { status: 'api_failed_final' };
}

// ============================================================
// main
// ============================================================
async function main() {
  console.log('=== 楽天ROOM自動投稿開始（v6: XHR+fetch+URL三段判定）===');
  console.log('EM:',  CONFIG.EMAIL ? '設定済み' : '未設定');
  console.log('PW:',  CONFIG.PASS  ? '設定済み' : '未設定');
  console.log('SID:', CONFIG.SHEET ? '設定済み' : '未設定');
  console.log('GSA:', process.env.GSA ? '設定済み' : '未設定');

  const items = await getUnpostedItems();
  console.log(`未投稿件数: ${items.length}件`);
  if (!items.length) { console.log('未投稿の商品がありません'); return; }

  const targets = items.slice(0, CONFIG.MAX_ITEMS_PER_RUN);
  const summary = {
    success: 0, manual_check: 0, api_failed: 0, api_failed_final: 0,
    login_failed: 0, form_not_found: 0,
    invalid_item_code: 0, empty_post_text: 0, exception: 0,
  };

  for (const item of targets) {
    console.log(`\n--- 投稿開始: ${(item.itemName || '').substring(0, 40)} ---`);
    console.log(`itemCode(C列): ${item.itemCode}`);

    const driver = createDriver();
    let result = { status: 'unknown' };
    try {
      result = await postItem(driver, item);
    } catch (e) {
      console.error('例外:', e.message);
      await screenshot(driver, `error_row${item.rowIndex}.png`);
      result = { status: 'exception', error: e.message };
    } finally {
      await driver.quit().catch(() => {});
    }

    summary[result.status] = (summary[result.status] || 0) + 1;

    if (result.status === 'success') {
      await updateStatus(item.rowIndex, '済');
    } else if (result.status === 'manual_check') {
      // API応答・URL変化のどちらも取れなかったケース
      // 実際は投稿されている可能性が高いので、重複防止のため「要確認」で記録
      await updateStatus(item.rowIndex, '要確認');
    } else if (
      result.status === 'invalid_item_code' ||
      result.status === 'empty_post_text'
    ) {
      await updateStatus(item.rowIndex, 'スキップ');
    }

    await sleep(CONFIG.WAIT_BETWEEN_POSTS_MS);
  }

  console.log('\n=== 実行サマリ ===');
  console.log(JSON.stringify(summary, null, 2));
  console.log('=== 楽天ROOM自動投稿完了 ===');
}

main().catch(e => { console.error('致命的:', e); process.exit(1); });
