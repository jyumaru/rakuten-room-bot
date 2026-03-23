const puppeteer = require('puppeteer');
const { google } = require('googleapis');

// ============================================================
// 設定
// ============================================================
const CONFIG = {
  RAKUTEN_EMAIL:    process.env.RAKUTEN_EMAIL,
  RAKUTEN_PASSWORD: process.env.RAKUTEN_PASSWORD,
  SPREADSHEET_ID:   process.env.SPREADSHEET_ID,
};

// ============================================================
// Googleスプレッドシートから未投稿の商品を取得
// ============================================================
async function getUnpostedItems() {
  const b64 = process.env.GOOGLE_SERVICE_ACCOUNT;
  const json = Buffer.from(b64, 'base64').toString('utf8');
  const credentials = JSON.parse(json);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: '楽天ROOM投稿リスト!A:H',
  });

  const rows = res.data.values || [];
  const items = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const status = row[7] || '未投稿';
    if (status === '未投稿') {
      items.push({
        rowIndex: i + 1,
        date:      row[0],
        category:  row[1],
        itemCode:  row[2],
        itemName:  row[3],
        price:     row[4],
        itemUrl:   row[5],
        postText:  row[6],
      });
    }
  }

  console.log(`未投稿件数: ${items.length}件`);
  return items;
}

// ============================================================
// 投稿済みに更新
// ============================================================
async function markAsPosted(rowIndex) {
  const b64 = process.env.GOOGLE_SERVICE_ACCOUNT;
  const json = Buffer.from(b64, 'base64').toString('utf8');
  const credentials = JSON.parse(json);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.update({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: `楽天ROOM投稿リスト!H${rowIndex}`,
    valueInputOption: 'RAW',
    requestBody: { values: [['済']] },
  });

  console.log(`行${rowIndex}を「済」に更新しました`);
}

// ============================================================
// 楽天ROOMに投稿（お気に入り経由）
// ============================================================
async function postToRakutenRoom(item) {
  const browser = await puppeteer.launch({
    headless: true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
    ],
  });

  const page = await browser.newPage();
  page.setDefaultTimeout(30000);

  try {
    // ① 楽天にログイン
    console.log('楽天にログイン中...');
    await page.goto('https://grp01.id.rakuten.co.jp/rms/nid/login', {
      waitUntil: 'networkidle2'
    });

    try {
      await page.waitForSelector('input[name="u"]', { timeout: 5000 });
      await page.type('input[name="u"]', CONFIG.RAKUTEN_EMAIL);
      await page.type('input[name="p"]', CONFIG.RAKUTEN_PASSWORD);
      await Promise.all([
        page.waitForNavigation({ waitUntil: 'networkidle2' }),
        page.click('input[type="submit"]'),
      ]);
      console.log('ログイン完了');
      // ログイン後に少し待機
      await new Promise(r => setTimeout(r, 3000));
    } catch (e) {
      console.log('ログインフォームなし（既にログイン済み）');
    }

    // ② 商品ページを開く
    console.log(`商品ページを開く: ${item.itemName.substring(0, 30)}`);
    await page.goto(item.itemUrl, { waitUntil: 'networkidle2' });
    await new Promise(r => setTimeout(r, 2000));

    // スクリーンショット①（商品ページ）
    await page.screenshot({ path: 'debug1.png', fullPage: false });
    console.log('商品ページのスクリーンショット保存');

    // クーポンポップアップを閉じる
    try {
      const closeBtn = await page.$('[hint="ポップアップを閉じる"]');
      if (closeBtn) {
        await closeBtn.click();
        console.log('ポップアップを閉じました');
        await new Promise(r => setTimeout(r, 1000));
      }
    } catch (e) {}

    // ③ お気に入りボタンをクリック（複数のセレクタを試す）
    console.log('お気に入りボタンを探しています...');
    let bookmarkClicked = false;

    const bookmarkSelectors = [
      '.floatingBookmarkAreaWrapper',
      '[data-ratid="item_bookmark"]',
      '.bookmark-button',
      'a[href*="my.bookmark.rakuten"]',
      '.item-bookmark',
    ];

    for (const selector of bookmarkSelectors) {
      try {
        await page.waitForSelector(selector, { timeout: 3000 });
        await page.click(selector);
        console.log(`お気に入りボタンクリック成功: ${selector}`);
        bookmarkClicked = true;
        break;
      } catch (e) {
        console.log(`セレクタ未発見: ${selector}`);
      }
    }

    if (!bookmarkClicked) {
      // JavaScriptで全リンクを確認
      const links = await page.evaluate(() => {
        return Array.from(document.querySelectorAll('a')).map(a => ({
          href: a.href,
          text: a.textContent.trim().substring(0, 30),
          className: a.className
        })).filter(l => l.href.includes('bookmark') || l.text.includes('お気に入り'));
      });
      console.log('お気に入り関連リンク:', JSON.stringify(links.slice(0, 5)));

      await page.screenshot({ path: 'debug2.png', fullPage: true });
      console.log('お気に入りボタンが見つかりません。スキップします。');
      await browser.close();
      return false;
    }

    // ④ お気に入り一覧ページに遷移
    await page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 10000 });
    console.log('お気に入り一覧URL:', page.url());
    await new Promise(r => setTimeout(r, 2000));

    // スクリーンショット②（お気に入り一覧）
    await page.screenshot({ path: 'debug2.png', fullPage: true });
    console.log('お気に入り一覧のスクリーンショット保存');

    // ⑤ 「ROOMに投稿」を探してクリック
    // ボタン・リンク・クリッカブル要素を全て出力
const allElements = await page.evaluate(() => {
  const elements = Array.from(document.querySelectorAll('a, button, [onclick], [data-*]'));
  return elements
    .map(el => ({
      tag: el.tagName,
      text: el.textContent.trim().substring(0, 50),
      href: el.href || '',
      onclick: el.getAttribute('onclick') || '',
      dataAttrs: Object.keys(el.dataset).join(','),
      className: el.className.substring(0, 50)
    }))
    .filter(el => el.text.includes('ROOM') || el.text.includes('コレ') || el.text.includes('投稿'));
});
console.log('ROOM関連要素:', JSON.stringify(allElements, null, 2));
    console.log('ROOMに投稿を探しています...');
    // ページ上の全リンクをログ出力（デバッグ用）
const allLinks = await page.evaluate(() => {
  return Array.from(document.querySelectorAll('a'))
    .map(a => ({ text: a.textContent.trim().substring(0, 50), href: a.href.substring(0, 100) }))
    .filter(l => l.text.length > 0);
});
console.log('ページ上のリンク一覧:');
allLinks.forEach(l => console.log(`  [${l.text}] → ${l.href}`));
    try {
      await page.waitForSelector('a[href*="room.rakuten.co.jp/mix/collect"]', {
        timeout: 10000
      });
      await page.click('a[href*="room.rakuten.co.jp/mix/collect"]');
      await page.waitForNavigation({ waitUntil: 'networkidle2' });
      console.log('ROOMに投稿フォームURL:', page.url());
    } catch (e) {
      console.log('ROOMに投稿ボタンが見つかりません');
      await page.screenshot({ path: 'debug3.png', fullPage: true });
      await browser.close();
      return false;
    }

    // ⑥ 投稿フォームに紹介文を入力
    console.log('紹介文を入力中...');
    try {
      await page.waitForSelector('#collect-content', { timeout: 10000 });
      await page.click('#collect-content');
      await page.evaluate((text) => {
        document.querySelector('#collect-content').value = text;
      }, item.postText);
    } catch (e) {
      console.log('投稿フォームが見つかりません');
      await page.screenshot({ path: 'debug3.png', fullPage: true });
      await browser.close();
      return false;
    }

    // ⑦ 投稿ボタンをクリック
    console.log('投稿中...');
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'networkidle2' }),
      page.click('button[type="submit"]'),
    ]);

    // ⑧ 投稿完了確認
    const url = page.url();
    console.log('投稿後URL:', url);
    await page.screenshot({ path: 'debug_final.png', fullPage: false });

    if (url.includes('complete') || url.includes('room.rakuten.co.jp')) {
      console.log(`✅ 投稿成功: ${item.itemName.substring(0, 40)}`);
      await browser.close();
      return true;
    } else {
      console.log(`❌ 投稿失敗: ${url}`);
      await browser.close();
      return false;
    }

  } catch (e) {
    console.error(`エラー: ${e.message}`);
    await page.screenshot({ path: 'debug_error.png', fullPage: true }).catch(() => {});
    await browser.close();
    return false;
  }
}

// ============================================================
// メイン処理
// ============================================================
async function main() {
  console.log('=== 楽天ROOM自動投稿開始 ===');

  const items = await getUnpostedItems();

  if (items.length === 0) {
    console.log('未投稿の商品がありません');
    return;
  }

  // 1回の実行で最大3件投稿（負荷軽減）
  const targetItems = items.slice(0, 3);

  for (const item of targetItems) {
    console.log(`\n--- 投稿開始: ${item.itemName.substring(0, 40)} ---`);

    const success = await postToRakutenRoom(item);

    if (success) {
      await markAsPosted(item.rowIndex);
    }

    // 投稿間隔を空ける（連続投稿対策）
    await new Promise(r => setTimeout(r, 5000));
  }

  console.log('\n=== 楽天ROOM自動投稿完了 ===');
}

main().catch(console.error);
