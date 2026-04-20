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
  MAX_ITEMS_PER_RUN: 5,        // 1回の実行での最大投稿数
  WAIT_BETWEEN_POSTS_MS: 5000, // 投稿間の待機時間
};

// ============================================================
// 禁止キーワードリスト（楽天ROOM API R145 対策）
//   投稿前に一括除去。API側で再度400が出た場合は message から
//   追加の禁止語を抽出してサニタイズしてリトライする。
// ============================================================
const BANNED_KEYWORDS = [
  // 医療・健康効果（薬機法抵触リスク）
  '腱鞘炎', '肩こり', '腰痛', '頭痛', '眼精疲労', '疲労回復', '不眠',
  '冷え性', 'むくみ', 'ダイエット効果', 'うつ',
  '治る', '治療', '改善', '予防', '効果', '効能', '症状',
  '医療', '医薬', '診断', '処方',
  // その他ROOMでよく引っかかるもの
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

// APIのエラーメッセージから禁止語を抽出（例: "禁止されたキーワードが含まれます: 腱鞘炎"）
function extractBannedFromApiMessage(message) {
  if (!message) return [];
  const m = message.match(/禁止されたキーワードが含まれます[:：]\s*(.+?)(?:$|。|\s)/);
  if (!m) return [];
  return m[1].split(/[、,・\s]+/).filter(Boolean);
}

// ============================================================
// Googleスプレッドシートから未投稿の商品を取得
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
      range: '楽天ROOM投稿リスト!A:H',
    });

    const rows = res.data.values || [];
    const items = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const status = row[7] || ''; // H列: ステータス
      if (status === '済') continue;

      items.push({
        rowIndex: i + 1,
        itemUrl:   row[0] || '', // A: 商品URL
        itemName:  row[1] || '', // B: 商品名
        price:     row[2] || '', // C: 価格
        shopName:  row[3] || '', // D: 店舗名
        postText:  row[4] || '', // E: 紹介文
        hashtags:  row[5] || '', // F: ハッシュタグ
        category:  row[6] || '', // G: カテゴリ
      });
    }
    return items;
  } catch (e) {
    console.error('スプレッドシート取得エラー:', e.message);
    return [];
  }
}

async function markAsPosted(rowIndex, statusText = '済') {
  try {
    const sheets = await getSheetsClient();
    await sheets.spreadsheets.values.update({
      spreadsheetId: CONFIG.SHEET,
      range: `楽天ROOM投稿リスト!H${rowIndex}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[statusText]] },
    });
    console.log(`行${rowIndex}を「${statusText}」に更新しました`);
  } catch (e) {
    console.error(`行${rowIndex}の更新エラー:`, e.message);
  }
}

// ============================================================
// Seleniumドライバー
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
  } catch (e) { /* ignore */ }
}

// ============================================================
// ログイン処理
// ============================================================
async function performLogin(driver) {
  console.log('ログイン処理開始...');

  // メール入力
  let emailInput = null;
  const emailSelectors = [
    By.name('u'), By.name('username'), By.name('email'),
    By.css('input[type="email"]'), By.css('input[type="text"]'),
  ];
  for (let attempt = 0; attempt < 5 && !emailInput; attempt++) {
    for (const sel of emailSelectors) {
      try {
        await driver.wait(until.elementLocated(sel), 2000);
        emailInput = await driver.findElement(sel);
        if (emailInput) break;
      } catch (e) {}
    }
    if (!emailInput) await sleep(1500);
  }
  if (!emailInput) {
    console.log('メール入力欄が見つかりません');
    return false;
  }
  await emailInput.clear();
  await emailInput.sendKeys(CONFIG.EMAIL);
  await sleep(800);

  // 次へボタン or Enter
  try {
    const nextBtn = await driver.findElement(By.css('button[type="submit"], input[type="submit"]'));
    await driver.executeScript('arguments[0].click();', nextBtn);
  } catch (e) {
    await emailInput.sendKeys(Key.RETURN);
  }
  await sleep(3000);

  // パスワード入力
  let passInput = null;
  for (let attempt = 0; attempt < 5 && !passInput; attempt++) {
    try {
      passInput = await driver.findElement(By.css('input[type="password"]'));
    } catch (e) { await sleep(1500); }
  }
  if (!passInput) {
    console.log('パスワード入力欄が見つかりません');
    return false;
  }
  await passInput.clear();
  await passInput.sendKeys(CONFIG.PASS);
  await sleep(800);

  try {
    const submitBtn = await driver.findElement(By.css('button[type="submit"], input[type="submit"]'));
    await driver.executeScript('arguments[0].click();', submitBtn);
  } catch (e) {
    await passInput.sendKeys(Key.RETURN);
  }
  await sleep(5000);

  const url = await driver.getCurrentUrl();
  if (url.includes('my.bookmark.rakuten.co.jp') || url.includes('rakuten.co.jp')) {
    console.log(`✅ ログイン成功: ${url}`);
    return true;
  }
  console.log(`ログイン後URL（想定外）: ${url}`);
  return false;
}

// ============================================================
// ブックマーク（お気に入り）ボタンクリック＋検証
//   クリック後、ROOM画面への遷移 or モーダル出現で成功判定
// ============================================================
async function clickBookmarkWithVerify(driver) {
  const selectors = [
    '[data-ratid="item_bookmark"]',
    '[data-ratid*="bookmark"]',
    '.floatingBookmarkAreaWrapper button',
    '.itemBookmarkAreaWrapper button',
    'button[class*="Bookmark"]',
    'a[href*="my.bookmark.rakuten.co.jp/bookmark/register"]',
    '[aria-label*="お気に入り"]',
    '[aria-label*="ブックマーク"]',
    'button:has(svg[class*="bookmark"])',
  ];

  for (const sel of selectors) {
    let elements = [];
    try {
      elements = await driver.findElements(By.css(sel));
    } catch (e) { continue; }
    if (!elements.length) continue;

    for (const el of elements) {
      try {
        const displayed = await el.isDisplayed().catch(() => false);
        if (!displayed) continue;

        await driver.executeScript('arguments[0].scrollIntoView({block: "center"});', el);
        await sleep(400);
        await driver.executeScript('arguments[0].click();', el);

        // ROOMへの遷移 or モーダル出現を最大6秒まで待機
        const ok = await driver.wait(async () => {
          const url = await driver.getCurrentUrl();
          if (url.includes('room.rakuten.co.jp')) return true;
          const handles = await driver.getAllWindowHandles();
          if (handles.length > 1) return true;
          try {
            const popups = await driver.findElements(
              By.css('[class*="roomPopup"], [class*="RoomModal"], [class*="bookmark-popup"]')
            );
            if (popups.length > 0) return true;
          } catch (e) {}
          return false;
        }, 6000).then(() => true).catch(() => false);

        if (ok) {
          console.log(`✅ ブックマーク成功: ${sel}`);
          return true;
        }
      } catch (e) { /* 次の要素へ */ }
    }
  }
  return false;
}

// ROOMリンクを取得（新タブ or 現タブ両対応）
async function findRoomLink(driver) {
  // 新タブに遷移していれば切り替え
  const handles = await driver.getAllWindowHandles();
  if (handles.length > 1) {
    await driver.switchTo().window(handles[handles.length - 1]);
    await sleep(1500);
  }

  // 現在URLがすでにROOMなら直接返す
  const currentUrl = await driver.getCurrentUrl();
  if (currentUrl.includes('room.rakuten.co.jp')) return currentUrl;

  // ページ内のROOMリンクを探す
  try {
    const link = await driver.findElement(By.css('a[href*="room.rakuten.co.jp/mix"]'));
    return await link.getAttribute('href');
  } catch (e) {}
  return null;
}

// ============================================================
// ROOM投稿: collect() を直接呼んでAPIレスポンスを捕捉
// ============================================================
async function postToRoom(driver, roomUrl, postText) {
  await driver.get(roomUrl);
  await sleep(4000);

  // OKポップアップを閉じる（出ない場合もあるので失敗しても続行）
  for (let i = 0; i < 3; i++) {
    try {
      const okBtn = await driver.findElement(By.xpath('//button[normalize-space(text())="OK"]'));
      await driver.executeScript('arguments[0].click();', okBtn);
      console.log(`OKクリック ${i + 1}回目`);
      await sleep(800);
    } catch (e) { break; }
  }

  // itemCodeをURLから抽出
  const itemCodeMatch = roomUrl.match(/itemcode=([^&]+)/);
  const itemCode = itemCodeMatch ? decodeURIComponent(itemCodeMatch[1]) : '';
  console.log(`itemCode: ${itemCode}`);

  // 紹介文を投入＋collect()呼び出し＋fetchインターセプトでAPI応答を捕捉
  const apiResultRaw = await driver.executeAsyncScript(`
    const callback = arguments[arguments.length - 1];
    const content = arguments[0];
    const itemCode = arguments[1];

    // fetchをフックして /api/collect のレスポンスを横取り
    const origFetch = window.fetch;
    let captured = null;
    window.fetch = async function(url, opts) {
      const res = await origFetch.apply(this, arguments);
      try {
        if (typeof url === 'string' && url.includes('/api/collect')) {
          const clone = res.clone();
          const text = await clone.text();
          captured = { status: res.status, body: text };
        }
      } catch (e) {}
      return res;
    };

    // textarea に値セット＋input/changeイベント発火
    const ta = document.querySelector('#collect-content');
    if (ta) {
      ta.value = content;
      ta.dispatchEvent(new Event('input',  { bubbles: true }));
      ta.dispatchEvent(new Event('change', { bubbles: true }));
    }

    // AngularJSスコープを取得してcollect()を呼ぶ
    try {
      const el = document.querySelector('[ng-controller], [data-ng-controller]') || document.body;
      const scope = angular.element(el).scope();
      if (scope && typeof scope.collect === 'function') {
        if (itemCode && !scope.itemCode) scope.itemCode = itemCode;
        if (scope.collectContent !== undefined) scope.collectContent = content;
        scope.$apply(function() { scope.collect(); });
      }
    } catch (e) {
      callback({ status: 'error', error: 'scope:' + e.message });
      return;
    }

    // APIレスポンスを最大15秒待機
    const start = Date.now();
    const timer = setInterval(function() {
      if (captured || Date.now() - start > 15000) {
        clearInterval(timer);
        callback(captured || { status: 'timeout' });
      }
    }, 300);
  `, postText, itemCode);

  return apiResultRaw;
}

function parseApiResult(raw) {
  if (!raw) return { ok: false, reason: 'no response' };
  if (raw.status === 'timeout') return { ok: false, reason: 'timeout' };
  if (raw.status === 'error')   return { ok: false, reason: raw.error || 'script error' };

  const httpStatus = raw.status;
  let body = {};
  try { body = JSON.parse(raw.body || '{}'); } catch (e) {}

  if (httpStatus >= 200 && httpStatus < 300 && body.status !== 'error') {
    return { ok: true, httpStatus, body };
  }
  return { ok: false, httpStatus, body, message: body.message || '' };
}

// ============================================================
// 1商品を投稿するメインフロー（リトライ制御つき）
// ============================================================
async function postItem(driver, item) {
  // 紹介文を事前サニタイズ
  let { cleaned, hits } = sanitizeContent(item.postText);
  if (hits.length) {
    console.log(`⚠️ 事前サニタイズ: ${hits.join(', ')} を除去`);
  }

  // 楽天商品ページへ
  console.log('楽天にログイン中...');
  await driver.get('https://my.bookmark.rakuten.co.jp/');
  await sleep(2000);
  const loggedIn = await performLogin(driver);
  if (!loggedIn) return { status: 'login_failed' };

  console.log(`商品ページを開く: ${item.itemName.substring(0, 40)}`);
  await driver.get(item.itemUrl);
  await sleep(4000);

  console.log('お気に入りボタンをクリック中...');
  const bookmarked = await clickBookmarkWithVerify(driver);
  if (!bookmarked) {
    console.log('❌ お気に入りボタンが見つかりません（スキップ）');
    return { status: 'bookmark_failed' };
  }

  const roomUrl = await findRoomLink(driver);
  if (!roomUrl) {
    console.log('❌ ROOMリンクが見つかりません');
    return { status: 'room_link_not_found' };
  }
  console.log(`ROOMリンク発見: ${roomUrl}`);

  // 最大2回トライ（初回失敗時、API messageから追加禁止語を抽出して再投稿）
  for (let attempt = 1; attempt <= 2; attempt++) {
    console.log(`紹介文を入力中... (試行 ${attempt}/2)`);
    const raw = await postToRoom(driver, roomUrl, cleaned);
    console.log(`APIレスポンス: ${JSON.stringify(raw).substring(0, 200)}`);

    const result = parseApiResult(raw);
    if (result.ok) {
      console.log(`✅ 投稿成功: ${item.itemName.substring(0, 40)}`);
      return { status: 'success' };
    }

    console.log(`❌ API失敗 http=${result.httpStatus}: ${result.message}`);

    // 禁止キーワードエラーなら抽出してサニタイズして再投稿
    if (result.httpStatus === 400 && /禁止されたキーワード/.test(result.message || '')) {
      const extra = extractBannedFromApiMessage(result.message);
      if (extra.length) {
        console.log(`⚠️ API指摘の禁止語を除去: ${extra.join(', ')}`);
        for (const w of extra) cleaned = cleaned.split(w).join('');
        cleaned = cleaned.replace(/[ 　]{2,}/g, ' ').trim();
        continue; // リトライ
      }
    }
    return { status: 'api_failed', detail: result };
  }
  return { status: 'api_failed_final' };
}

// ============================================================
// main
// ============================================================
async function main() {
  console.log('=== 楽天ROOM自動投稿開始（Selenium版）===');
  console.log('EM:',  CONFIG.EMAIL ? '設定済み' : '未設定');
  console.log('PW:',  CONFIG.PASS  ? '設定済み' : '未設定');
  console.log('SID:', CONFIG.SHEET ? '設定済み' : '未設定');
  console.log('GSA:', process.env.GSA ? '設定済み' : '未設定');

  const items = await getUnpostedItems();
  console.log(`未投稿件数: ${items.length}件`);
  if (!items.length) {
    console.log('未投稿の商品がありません');
    return;
  }

  const targets = items.slice(0, CONFIG.MAX_ITEMS_PER_RUN);
  const summary = { success: 0, bookmark_failed: 0, room_link_not_found: 0, api_failed: 0, login_failed: 0 };

  for (const item of targets) {
    console.log(`\n--- 投稿開始: ${item.itemName.substring(0, 40)} ---`);

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

    // 成功したシートに記録。bookmark/api失敗は後で手動対応できるよう別ステータスで残す
    if (result.status === 'success') {
      await markAsPosted(item.rowIndex, '済');
    } else if (result.status === 'bookmark_failed' || result.status === 'room_link_not_found') {
      await markAsPosted(item.rowIndex, 'スキップ');
    }
    // api_failed はステータス更新しない（次回再挑戦させる）

    await sleep(CONFIG.WAIT_BETWEEN_POSTS_MS);
  }

  console.log('\n=== 実行サマリ ===');
  console.log(JSON.stringify(summary, null, 2));
  console.log('=== 楽天ROOM自動投稿完了 ===');
}

main().catch((e) => {
  console.error('致命的エラー:', e);
  process.exit(1);
});
