const puppeteer = require('puppeteer');
const { google } = require('googleapis');

// ============================================================
// 設定
// ============================================================
const CONFIG = {
  EMAIL: process.env.EM,
  PASS:  process.env.PW,
  SHEET: process.env.SID,
};

// ============================================================
// Googleスプレッドシートから未投稿の商品を取得
// ============================================================
async function getUnpostedItems() {
  try {
    const b64 = process.env.GSA;
    console.log('GSA長さ:', b64 ? b64.length : 'null');
    const json = Buffer.from(b64, 'base64').toString('utf8');
    console.log('JSON先頭:', json.substring(0, 30));
    const credentials = JSON.parse(json);
    console.log('認証情報タイプ:', credentials.type);

    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.SHEET,
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

  } catch (e) {
    console.error('getUnpostedItemsエラー:', e.message);
    throw e;
  }
}

// ============================================================
// 投稿済みに更新
// ============================================================
async function markAsPosted(rowIndex) {
  try {
    const b64 = process.env.GSA;
    const json = Buffer.from(b64, 'base64').toString('utf8');
    const credentials = JSON.parse(json);
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });
    await sheets.spreadsheets.values.update({
      spreadsheetId: CONFIG.SHEET,
      range: `楽天ROOM投稿リスト!H${rowIndex}`,
      valueInputOption: 'RAW',
      requestBody: { values: [['済']] },
    });

    console.log(`行${rowIndex}を「済」に更新しました`);

  } catch (e) {
    console.error('markAsPostedエラー:', e.message);
  }
}

// ============================================================
// 楽天ROOMに投稿（お気に入りブックマーク経由）
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
    await page.goto(
      'https://login.account.rakuten.com/sso/authorize?client_id=rakuten_ichiba_web&redirect_uri=https://www.rakuten.co.jp/&response_type=code&scope=openid',
      { waitUntil: 'networkidle2' }
    );

    await page.screenshot({ path: 'debug_login.png', fullPage: false });
    console.log('ログインページURL:', page.url());

    try {
      await page.waitForSelector('input[name="username"], input[type="email"], #email', { timeout: 8000 });
      const emailInput =
        await page.$('input[name="username"]') ||
        await page.$('input[type="email"]') ||
        await page.$('#email');
      await emailInput.type(CONFIG.EMAIL);

      const nextBtn = await page.$('button[type="submit"]');
      if (nextBtn) {
        await Promise.all([
          page.waitForNavigation({ waitUntil: 'networkidle2' }).catch(() => {}),
          nextBtn.click(),
        ]);
        await new Promise(r => setTimeout(r, 2000));
      }

      await page.screenshot({ path: 'debug_login2.png', fullPage: false });

      const passInput = await page.$('input[type="password"]');
      if (passInput) {
        await passInput.type(CONFIG.PASS);
        await Promise.all([
          page.waitForNavigation({ waitUntil: 'networkidle2' }),
          page.click('button[type="submit"]'),
        ]);
      }

      console.log('ログイン完了');
      await new Promise(r => setTimeout(r, 3000));
      console.log('ログイン後URL:', page.url());
      await page.screenshot({ path: 'debug_login3.png', fullPage: false });

    } catch (e) {
      console.log('ログインエラー:', e.message);
      await page.screenshot({ path: 'debug_login_error.png', fullPage: false });
    }

    // ② 商品ページを開く
    console.log(`商品ページを開く: ${item.itemName.substring(0, 30)}`);
    await page.goto(item.itemUrl, { waitUntil: 'networkidle2' });
    await new Promise(r => setTimeout(r, 2000));
    await page.screenshot({ path: 'debug1.png', fullPage: false });

    // クーポンポップアップを閉じる
    try {
      const closeBtn = await page.$('[hint="ポップアップを閉じる"]');
      if (closeBtn) {
        await closeBtn.click();
        await new Promise(r => setTimeout(r, 1000));
        console.log('ポップアップを閉じました');
      }
    } catch (e) {}

    // ③ お気に入りボタンをクリック
    console.log('お気に入りボタンをクリック中...');
    let bookmarkClicked = false;

    const bookmarkSelectors = [
      '.floatingBookmarkAreaWrapper',
      '[data-ratid="item_bookmark"]',
      'a[href*="my.bookmark.rakuten"]',
      '.bookmark-button',
      'button[class*="bookmark"]',
      'a[class*="bookmark"]',
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
      console.log('お気に入りボタンが見つかりません。スキップします。');
      await page.screenshot({ path: 'debug_no_bookmark.png', fullPage: true });
      await browser.close();
      return false;
    }

    // ④ お気に入りブックマークページへの遷移を待つ
    await page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 20000 }).catch(e => {
      console.log('ナビゲーションタイムアウト（続行）:', e.message);
    });
    await new Promise(r => setTimeout(r, 3000));
    const bookmarkUrl = page.url();
    console.log('遷移先URL:', bookmarkUrl);
    await page.screenshot({ path: 'debug2.png', fullPage: true });

    // ログインページに飛ばされた場合は再ログイン
    if (bookmarkUrl.includes('login') || bookmarkUrl.includes('sign_in')) {
      console.log('ログインページにリダイレクトされました。再ログインします...');
      try {
        const emailInput =
          await page.$('input[name="username"]') ||
          await page.$('input[type="email"]') ||
          await page.$('#email');
        if (emailInput) await emailInput.type(CONFIG.EMAIL);

        const nextBtn = await page.$('button[type="submit"]');
        if (nextBtn) {
          await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle2' }).catch(() => {}),
            nextBtn.click(),
          ]);
          await new Promise(r => setTimeout(r, 2000));
        }

        const passInput = await page.$('input[type="password"]');
        if (passInput) {
          await passInput.type(CONFIG.PASS);
          await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle2' }),
            page.click('button[type="submit"]'),
          ]);
        }

        console.log('再ログイン完了');
        await new Promise(r => setTimeout(r, 3000));
        console.log('再ログイン後URL:', page.url());
        await page.screenshot({ path: 'debug2b.png', fullPage: true });
      } catch (e) {
        console.log('再ログイン失敗:', e.message);
        await browser.close();
        return false;
      }
    }

    // ⑤ ROOM関連要素をログ出力
    console.log('ROOM関連要素を検索中...');
    const allElements = await page.evaluate(() => {
      const elements = Array.from(document.querySelectorAll('a, button, [onclick]'));
      return elements
        .map(el => ({
          tag: el.tagName,
          text: el.textContent.trim().substring(0, 50),
          href: el.href || '',
          onclick: el.getAttribute('onclick') || '',
          className: el.className.substring(0, 50)
        }))
        .filter(el =>
          el.text.includes('ROOM') ||
          el.text.includes('コレ') ||
          el.text.includes('投稿') ||
          el.href.includes('room.rakuten')
        );
    });
    console.log('ROOM関連要素:', JSON.stringify(allElements, null, 2));

    // ⑥ 「ROOMに投稿」リンクをクリック
    console.log('ROOMに投稿リンクをクリック中...');
    try {
      const clicked = await page.evaluate(() => {
        const links = Array.from(document.querySelectorAll('a, button, [onclick]'));
        const roomLink = links.find(a =>
          (a.href && a.href.includes('room.rakuten.co.jp/mix/collect')) ||
          a.textContent.includes('ROOMに投稿') ||
          a.textContent.includes('コレ！')
        );
        if (roomLink) {
          roomLink.click();
          return roomLink.href || roomLink.textContent;
        }
        return null;
      });

      if (clicked) {
        console.log('ROOMリンククリック成功:', clicked);
        await page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }).catch(() => {});
        await new Promise(r => setTimeout(r, 2000));
      } else {
        console.log('ROOMに投稿リンクが見つかりません');
        await page.screenshot({ path: 'debug3.png', fullPage: true });
        await browser.close();
        return false;
      }
    } catch (e) {
      console.log('ROOMリンククリックエラー:', e.message);
      await browser.close();
      return false;
    }

    console.log('ROOMコレクトフォームURL:', page.url());
    await page.screenshot({ path: 'debug3.png', fullPage: true });

    // ⑦ 投稿フォームに紹介文を入力
    console.log('紹介文を入力中...');
    try {
      await page.waitForSelector('#collect-content', { timeout: 10000 });
      await page.evaluate(() => {
        document.querySelector('#collect-content').value = '';
      });
      await page.evaluate((text) => {
        document.querySelector('#collect-content').value = text;
      }, item.postText);
      console.log('紹介文の入力完了');
    } catch (e) {
      console.log('投稿フォームが見つかりません:', page.url());
      await page.screenshot({ path: 'debug_no_form.png', fullPage: true });
      await browser.close();
      return false;
    }

    await page.screenshot({ path: 'debug4.png', fullPage: false });

    // ⑧ 「完了」ボタンをクリック
    console.log('「完了」ボタンをクリック中...');
    try {
      await Promise.all([
        page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }),
        page.click('button[type="submit"]'),
      ]);
    } catch (e) {
      console.log('ナビゲーションなしで投稿完了の可能性:', e.message);
    }

    // ⑨ 投稿完了確認
    const finalUrl = page.url();
    console.log('投稿後URL:', finalUrl);
    await page.screenshot({ path: 'debug_final.png', fullPage: false });

    if (
      finalUrl.includes('complete') ||
      finalUrl.includes('room.rakuten.co.jp') ||
      finalUrl.includes('collect')
    ) {
      console.log(`✅ 投稿成功: ${item.itemName.substring(0, 40)}`);
      await browser.close();
      return true;
    } else {
      console.log(`❌ 投稿失敗: ${finalUrl}`);
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

  // 環境変数チェック
  console.log('EM:', process.env.EM ? '設定済み' : '未設定');
  console.log('PW:', process.env.PW ? '設定済み' : '未設定');
  console.log('SID:', process.env.SID ? '設定済み' : '未設定');
  console.log('GSA:', process.env.GSA ? '設定済み' : '未設定');

  const items = await getUnpostedItems();

  if (items.length === 0) {
    console.log('未投稿の商品がありません');
    return;
  }

  const targetItems = items.slice(0, 3);

  for (const item of targetItems) {
    console.log(`\n--- 投稿開始: ${item.itemName.substring(0, 40)} ---`);

    const success = await postToRakutenRoom(item);

    if (success) {
      await markAsPosted(item.rowIndex);
    }

    await new Promise(r => setTimeout(r, 5000));
  }

  console.log('\n=== 楽天ROOM自動投稿完了 ===');
}

main().catch(console.error);
