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

// スプレッドシート列マッピング（GAS setupHeader と完全一致）
const COL = {
  CREATED_AT: 0, CATEGORY: 1, ITEM_CODE: 2, ITEM_NAME: 3,
  PRICE: 4, ITEM_URL: 5, POST_TEXT: 6, STATUS: 7,
};

// 禁止キーワード
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

// ============================================================
// アフィリエイトURLを展開して素のitem.rakuten.co.jpのURLを取り出す
// ============================================================
function unwrapAffiliateUrl(url) {
  if (!url) return url;
  try {
    // hb.afl.rakuten.co.jp のリダイレクタ → pc パラメータに元URL
    if (url.includes('hb.afl.rakuten.co.jp')) {
      const u = new URL(url);
      const pc = u.searchParams.get('pc');
      if (pc) {
        const decoded = decodeURIComponent(pc);
        if (/^https?:\/\//.test(decoded)) return decoded;
      }
    }
    return url;
  } catch (e) {
    return url;
  }
}

// item.rakuten.co.jp/{shop}/{itemId}/ から shopCode:itemId を取り出す
function parseShopAndItem(url) {
  try {
    const u = new URL(url);
    if (!u.hostname.includes('item.rakuten.co.jp')) return null;
    const parts = u.pathname.split('/').filter(Boolean);
    if (parts.length < 2) return null;
    return { shop: parts[0], item: parts[1] };
  } catch (e) {
    return null;
  }
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

      const itemUrl = row[COL.ITEM_URL] || '';
      if (!itemUrl) continue;

      items.push({
        rowIndex:  i + 1,
        createdAt: row[COL.CREATED_AT] || '',
        category:  row[COL.CATEGORY]   || '',
        itemCode:  row[COL.ITEM_CODE]  || '',
        itemName:  row[COL.ITEM_NAME]  || '',
        price:     row[COL.PRICE]      || '',
        itemUrl:   itemUrl,
        postText:  row[COL.POST_TEXT]  || '',
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
  } catch (e) {}
}

async function waitForPageReady(driver, timeoutMs = 15000) {
  try {
    await driver.wait(async () => {
      const state = await driver.executeScript('return document.readyState');
      return state === 'complete';
    }, timeoutMs);
  } catch (e) {}
  // アフィリエイトリダイレクタの場合、読込完了しても最終ページじゃないことがあるので追加で待機
  await sleep(1500);
}

// ============================================================
// ログイン処理
// ============================================================
async function performLogin(driver) {
  console.log('ログイン処理開始...');
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
  if (!emailInput) { console.log('メール入力欄なし'); return false; }

  await emailInput.clear();
  await emailInput.sendKeys(CONFIG.EMAIL);
  await sleep(800);
  try {
    const nextBtn = await driver.findElement(By.css('button[type="submit"], input[type="submit"]'));
    await driver.executeScript('arguments[0].click();', nextBtn);
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
    const submitBtn = await driver.findElement(By.css('button[type="submit"], input[type="submit"]'));
    await driver.executeScript('arguments[0].click();', submitBtn);
  } catch (e) { await passInput.sendKeys(Key.RETURN); }
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
// ブックマーク（お気に入り）ボタンクリック＋検証（強化版）
// ============================================================
async function clickBookmarkWithVerify(driver) {
  // 画面下部までスクロールしてDOMを全部ロードさせる（遅延読み込み対策）
  try {
    await driver.executeScript(`
      window.scrollTo(0, document.body.scrollHeight / 2);
    `);
    await sleep(800);
    await driver.executeScript(`window.scrollTo(0, 0);`);
    await sleep(500);
  } catch (e) {}

  const cssSelectors = [
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
  const xpathSelectors = [
    "//button[contains(., 'お気に入り')]",
    "//a[contains(., 'お気に入り')]",
    "//button[contains(., 'ブックマーク')]",
  ];

  const tryClick = async (el) => {
    const displayed = await el.isDisplayed().catch(() => false);
    if (!displayed) return false;

    await driver.executeScript('arguments[0].scrollIntoView({block: "center"});', el);
    await sleep(400);
    await driver.executeScript('arguments[0].click();', el);

    return await driver.wait(async () => {
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
  };

  // CSS セレクタで探す
  for (const sel of cssSelectors) {
    let els = [];
    try { els = await driver.findElements(By.css(sel)); } catch (e) { continue; }
    for (const el of els) {
      try {
        if (await tryClick(el)) {
          console.log(`✅ ブックマーク成功 (css): ${sel}`);
          return true;
        }
      } catch (e) {}
    }
  }
  // XPath セレクタで探す
  for (const xp of xpathSelectors) {
    let els = [];
    try { els = await driver.findElements(By.xpath(xp)); } catch (e) { continue; }
    for (const el of els) {
      try {
        if (await tryClick(el)) {
          console.log(`✅ ブックマーク成功 (xpath): ${xp}`);
          return true;
        }
      } catch (e) {}
    }
  }
  return false;
}

// ============================================================
// フォールバック: itemCode から直接ブックマーク登録URLを踏む
// ============================================================
async function bookmarkViaDirectUrl(driver, item) {
  // itemCode (C列) は "shop:itemId" 形式
  let shop = '', itemId = '';
  if (item.itemCode && item.itemCode.includes(':')) {
    [shop, itemId] = item.itemCode.split(':');
  } else {
    // itemCodeが無ければURLから推測
    const parsed = parseShopAndItem(item.resolvedUrl || item.itemUrl);
    if (parsed) { shop = parsed.shop; itemId = parsed.item; }
  }
  if (!shop || !itemId) return null;

  const bookmarkUrl =
    `https://my.bookmark.rakuten.co.jp/bookmark/register?shopCode=${encodeURIComponent(shop)}&itemCode=${encodeURIComponent(itemId)}`;
  console.log(`直接ブックマーク登録を試行: ${bookmarkUrl}`);
  try {
    await driver.get(bookmarkUrl);
    await waitForPageReady(driver);

    // ROOMへ遷移する or ROOMリンクが出てくるのを待つ
    for (let i = 0; i < 3; i++) {
      const currentUrl = await driver.getCurrentUrl();
      if (currentUrl.includes('room.rakuten.co.jp')) return currentUrl;
      try {
        const link = await driver.findElement(By.css('a[href*="room.rakuten.co.jp/mix"]'));
        const href = await link.getAttribute('href');
        if (href) return href;
      } catch (e) {}
      await sleep(1500);
    }
  } catch (e) {
    console.log(`直接ブックマーク失敗: ${e.message}`);
  }
  return null;
}

async function findRoomLink(driver) {
  const handles = await driver.getAllWindowHandles();
  if (handles.length > 1) {
    await driver.switchTo().window(handles[handles.length - 1]);
    await sleep(1500);
  }
  const currentUrl = await driver.getCurrentUrl();
  if (currentUrl.includes('room.rakuten.co.jp')) return currentUrl;

  try {
    const link = await driver.findElement(By.css('a[href*="room.rakuten.co.jp/mix"]'));
    return await link.getAttribute('href');
  } catch (e) {}
  return null;
}

// ============================================================
// ROOM投稿
// ============================================================
async function postToRoom(driver, roomUrl, postText) {
  await driver.get(roomUrl);
  await sleep(4000);

  for (let i = 0; i < 3; i++) {
    try {
      const okBtn = await driver.findElement(By.xpath('//button[normalize-space(text())="OK"]'));
      await driver.executeScript('arguments[0].click();', okBtn);
      console.log(`OKクリック ${i + 1}回目`);
      await sleep(800);
    } catch (e) { break; }
  }

  const itemCodeMatch = roomUrl.match(/itemcode=([^&]+)/);
  const itemCode = itemCodeMatch ? decodeURIComponent(itemCodeMatch[1]) : '';
  console.log(`itemCode: ${itemCode}`);

  const apiResultRaw = await driver.executeAsyncScript(`
    const callback = arguments[arguments.length - 1];
    const content = arguments[0];
    const itemCode = arguments[1];

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

    const ta = document.querySelector('#collect-content');
    if (ta) {
      ta.value = content;
      ta.dispatchEvent(new Event('input',  { bubbles: true }));
      ta.dispatchEvent(new Event('change', { bubbles: true }));
    }

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
// 1商品を投稿するメインフロー
// ============================================================
async function postItem(driver, item) {
  // アフィリエイトURLを展開
  item.resolvedUrl = unwrapAffiliateUrl(item.itemUrl);
  if (item.resolvedUrl !== item.itemUrl) {
    console.log(`アフィリエイトURL展開: ${item.resolvedUrl}`);
  }

  if (!/^https?:\/\//.test(item.resolvedUrl)) {
    console.log(`❌ 無効なURL: "${item.resolvedUrl}"`);
    return { status: 'invalid_url' };
  }

  let { cleaned, hits } = sanitizeContent(item.postText);
  if (hits.length) console.log(`⚠️ 事前サニタイズ: ${hits.join(', ')} を除去`);
  if (!cleaned) { console.log('❌ 紹介文が空'); return { status: 'empty_post_text' }; }

  console.log('楽天にログイン中...');
  await driver.get('https://my.bookmark.rakuten.co.jp/');
  await sleep(2000);
  const loggedIn = await performLogin(driver);
  if (!loggedIn) return { status: 'login_failed' };

  console.log(`商品ページを開く: ${item.itemName.substring(0, 40)}`);
  await driver.get(item.resolvedUrl);
  await waitForPageReady(driver);

  // ログが長くなるので現在URLも出しておく
  const landedUrl = await driver.getCurrentUrl();
  console.log(`現在URL: ${landedUrl.substring(0, 80)}...`);

  console.log('お気に入りボタンをクリック中...');
  let bookmarked = await clickBookmarkWithVerify(driver);

  let roomUrl = null;
  if (bookmarked) {
    roomUrl = await findRoomLink(driver);
  }

  // フォールバック: 直接ブックマーク登録URLを踏む
  if (!roomUrl) {
    console.log('⚠️ 通常のブックマーク失敗、フォールバック試行');
    roomUrl = await bookmarkViaDirectUrl(driver, item);
  }

  if (!roomUrl) {
    console.log('❌ ROOMリンク取得失敗');
    return { status: 'bookmark_failed' };
  }
  console.log(`ROOMリンク発見: ${roomUrl}`);

  for (let attempt = 1; attempt <= 2; attempt++) {
    console.log(`紹介文を入力中... (試行 ${attempt}/2)`);
    const raw = await postToRoom(driver, roomUrl, cleaned);
    const result = parseApiResult(raw);

    if (result.ok) {
      console.log(`✅ 投稿成功: ${item.itemName.substring(0, 40)}`);
      return { status: 'success' };
    }
    console.log(`❌ API失敗 http=${result.httpStatus}: ${result.message}`);

    if (result.httpStatus === 400 && /禁止されたキーワード/.test(result.message || '')) {
      const extra = extractBannedFromApiMessage(result.message);
      if (extra.length) {
        console.log(`⚠️ API指摘の禁止語を除去: ${extra.join(', ')}`);
        for (const w of extra) cleaned = cleaned.split(w).join('');
        cleaned = cleaned.replace(/[ 　]{2,}/g, ' ').trim();
        continue;
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
  if (!items.length) { console.log('未投稿の商品がありません'); return; }

  const targets = items.slice(0, CONFIG.MAX_ITEMS_PER_RUN);
  const summary = {
    success: 0, bookmark_failed: 0, room_link_not_found: 0,
    api_failed: 0, api_failed_final: 0, login_failed: 0,
    invalid_url: 0, empty_post_text: 0, exception: 0,
  };

  for (const item of targets) {
    console.log(`\n--- 投稿開始: ${(item.itemName || '').substring(0, 40)} ---`);

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
    } else if (
      result.status === 'invalid_url' ||
      result.status === 'empty_post_text'
    ) {
      await updateStatus(item.rowIndex, 'スキップ');
    }
    // bookmark_failed / api_failed 系はスキップ固定しない → 次回リトライ可能

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
