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
};

// ============================================================
// Googleスプレッドシートから未投稿の商品を取得
// ============================================================
async function getUnpostedItems() {
  try {
    const b64 = process.env.GSA;
    const json = Buffer.from(b64, 'base64').toString('utf8');
    const credentials = JSON.parse(json);

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
// Seleniumドライバーを作成
// ============================================================
function createDriver() {
  const options = new chrome.Options();
  options.addArguments('--headless');
  options.addArguments('--no-sandbox');
  options.addArguments('--disable-dev-shm-usage');
  options.addArguments('--disable-gpu');
  options.addArguments('--window-size=1280,800');
  options.addArguments('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');

  return new Builder()
    .forBrowser('chrome')
    .setChromeOptions(options)
    .build();
}

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function screenshot(driver, filename) {
  try {
    const img = await driver.takeScreenshot();
    fs.writeFileSync(filename, img, 'base64');
    console.log(`スクリーンショット: ${filename}`);
  } catch (e) {}
}

// ============================================================
// ログイン処理
// ============================================================
async function performLogin(driver, screenshotPrefix, successUrlContains) {
  console.log('ログイン処理開始...');
  await screenshot(driver, `${screenshotPrefix}_1.png`);

  try {
    let emailInput = null;
    const emailSelectors = [
      By.name('email'),
      By.name('username'),
      By.css('input[type="email"]'),
      By.css('input[type="text"]'),
    ];

    for (let attempt = 0; attempt < 3; attempt++) {
      for (const sel of emailSelectors) {
        try {
          emailInput = await driver.findElement(sel);
          if (emailInput) break;
        } catch (e) {}
      }
      if (emailInput) break;
      await sleep(2000);
    }

    if (!emailInput) {
      console.log('メール入力欄が見つかりません');
      return false;
    }

    console.log('メールアドレスを入力中...');
    await emailInput.clear();
    await emailInput.sendKeys(CONFIG.EMAIL);
    await sleep(1000);

    try {
      const btn = await driver.findElement(By.css('button[type="submit"]'));
      await driver.executeScript('arguments[0].click();', btn);
    } catch (e) {
      await emailInput.sendKeys(Key.RETURN);
    }
    await sleep(3000);
    await screenshot(driver, `${screenshotPrefix}_2.png`);

    const passInput = await driver.findElement(By.css('input[type="password"]'));
    console.log('パスワードを入力中...');
    await passInput.clear();

    for (const char of CONFIG.PASS) {
      await passInput.sendKeys(char);
      await sleep(50);
    }
    await sleep(1000);

    try {
      const btn = await driver.findElement(By.css('button[type="submit"]'));
      await driver.executeScript('arguments[0].click();', btn);
    } catch (e) {
      await passInput.sendKeys(Key.RETURN);
    }

    console.log(`リダイレクト待機中... (目標URL: ${successUrlContains})`);
    for (let i = 0; i < 15; i++) {
      await sleep(1000);
      const url = await driver.getCurrentUrl();
      if (url.includes(successUrlContains) && !url.includes('sign_in') && !url.includes('login')) {
        console.log(`✅ ログイン成功: ${url}`);
        await screenshot(driver, `${screenshotPrefix}_3.png`);
        return true;
      }
    }

    const finalUrl = await driver.getCurrentUrl();
    console.log('❌ ログイン失敗（タイムアウト）:', finalUrl);
    await screenshot(driver, `${screenshotPrefix}_failed.png`);
    return false;

  } catch (e) {
    console.log('ログインエラー:', e.message);
    await screenshot(driver, `${screenshotPrefix}_error.png`);
    return false;
  }
}

// ============================================================
// 楽天ROOMに投稿
// ============================================================
async function postToRakutenRoom(item) {
  const driver = await createDriver();

  try {
    // ① 楽天ブックマーク認証URLでログイン
    console.log('楽天にログイン中...');
    await driver.get('https://login.account.rakuten.com/sso/authorize?client_id=rakuten_ichiba_bookmark_web&redirect_uri=https://my.bookmark.rakuten.co.jp&response_type=code&scope=openid&state=login#/sign_in');
    await sleep(3000);

    const loggedIn = await performLogin(driver, 'debug_login', 'my.bookmark.rakuten.co.jp');
    if (!loggedIn) {
      console.log('ログイン失敗');
      await driver.quit();
      return false;
    }

    // ② 楽天トップページでセッション確立
    await driver.get('https://www.rakuten.co.jp/?l-id=pc_header_logo');
    await sleep(2000);

    // ③ 商品ページを開く
    console.log(`商品ページを開く: ${item.itemName.substring(0, 30)}`);
    await driver.get(item.itemUrl);
    await sleep(3000);
    await screenshot(driver, 'debug1.png');

    // クーポンポップアップを閉じる
    try {
      const closeBtn = await driver.findElement(By.css('[hint="ポップアップを閉じる"]'));
      await closeBtn.click();
      await sleep(1000);
    } catch (e) {}

    // ④ お気に入りボタンをクリック
    console.log('お気に入りボタンをクリック中...');
    let bookmarkClicked = false;

    for (const sel of [
      '[data-ratid="item_bookmark"]',
      '[data-ratid*="bookmark"]',
      '.floatingBookmarkAreaWrapper',
      '.itemBookmarkAreaWrapper',
      '[class*="Bookmark"]',
      '[class*="bookmark"]',
    ]) {
      try {
        const el = await driver.findElement(By.css(sel));
        await driver.executeScript('arguments[0].click();', el);
        console.log(`お気に入りボタンクリック成功: ${sel}`);
        bookmarkClicked = true;
        break;
      } catch (e) {
        console.log(`セレクタ未発見: ${sel}`);
      }
    }

    if (!bookmarkClicked) {
      console.log('お気に入りボタンが見つかりません');
      await screenshot(driver, 'debug_no_bookmark.png');
      await driver.quit();
      return false;
    }

    // ⑤ お気に入りページへのリダイレクトを待つ
    for (let i = 0; i < 15; i++) {
      await sleep(1000);
      const url = await driver.getCurrentUrl();
      if (url.includes('my.bookmark.rakuten.co.jp')) {
        console.log('✅ お気に入りページに到達');
        break;
      }
      if (url.includes('login') || url.includes('sign_in')) {
        await performLogin(driver, 'debug_relogin', 'my.bookmark.rakuten.co.jp');
        break;
      }
    }

    await screenshot(driver, 'debug2.png');

    // ⑥ ROOMに投稿リンクを取得
    await sleep(2000);
    const roomLink = await driver.executeScript(() => {
      const links = Array.from(document.querySelectorAll('a'));
      const link = links.find(a => a.href && a.href.includes('room.rakuten.co.jp'));
      return link ? link.href : null;
    });

    if (!roomLink) {
      console.log('ROOMリンクが見つかりません');
      await screenshot(driver, 'debug3.png');
      await driver.quit();
      return false;
    }

    console.log('ROOMリンク発見:', roomLink);

    let collectUrl = roomLink.includes('/mix?')
      ? roomLink.replace('/mix?', '/mix/collect?')
      : roomLink;
    console.log('投稿フォームURL:', collectUrl);

    // ⑦ ROOMの投稿フォームにアクセス
    await driver.get(collectUrl);
    await sleep(3000);

    const curUrl = await driver.getCurrentUrl();
    if (curUrl.includes('login') || curUrl.includes('sign_in')) {
      console.log('ROOMログインページ検出。ログイン中...');
      await performLogin(driver, 'debug_room_relogin', 'room.rakuten.co.jp');
      await sleep(3000);
      await driver.get(collectUrl);
      await sleep(3000);
    }

    console.log('ROOMコレクトフォームURL:', await driver.getCurrentUrl());
    await screenshot(driver, 'debug3.png');

    // ⑧ 投稿フォームに紹介文を入力
    console.log('紹介文を入力中...');
    try {
      await driver.wait(until.elementLocated(By.id('collect-content')), 10000);

      await driver.executeScript(`
        const el = document.querySelector('#collect-content');
        el.value = arguments[0];
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('keyup', { bubbles: true }));
      `, item.postText);

      const charCount = await driver.executeScript(() => {
        const el = document.querySelector('#collect-content');
        return el ? el.value.length : 0;
      });
      console.log(`紹介文入力完了（${charCount}文字）`);

    } catch (e) {
      console.log('投稿フォームが見つかりません:', await driver.getCurrentUrl());
      await screenshot(driver, 'debug_no_form.png');
      await driver.quit();
      return false;
    }

    await screenshot(driver, 'debug4.png');

    // ⑨ ページHTMLをファイルに保存（Shadow DOM調査用）
    const pageHtml = await driver.executeScript(() => document.documentElement.outerHTML);
    fs.writeFileSync('page_debug.html', pageHtml);
    console.log('ページHTML保存完了');

    // Shadow DOMを含む全ボタンを調査
    const shadowBtns = await driver.executeScript(() => {
      const results = [];
      function findButtons(root, depth) {
        const btns = root.querySelectorAll('button');
        btns.forEach(b => results.push({
          text: b.textContent.trim().substring(0, 30),
          className: b.className.substring(0, 50),
          type: b.type,
          depth: depth,
          inShadow: depth > 0
        }));
        root.querySelectorAll('*').forEach(el => {
          if (el.shadowRoot) findButtons(el.shadowRoot, depth + 1);
        });
      }
      findButtons(document, 0);
      return results;
    });
    console.log('Shadow DOM含む全ボタン:', JSON.stringify(shadowBtns, null, 2));

    // ⑩ OKポップアップを閉じる
    console.log('OKポップアップを閉じます...');
    for (let i = 0; i < 10; i++) {
      try {
        const okBtn = await driver.findElement(
          By.xpath('//button[normalize-space(text())="OK"]')
        );
        await driver.executeScript('arguments[0].click();', okBtn);
        console.log(`OKボタンをクリック（${i + 1}回目）`);
        await sleep(800);
      } catch (e) {
        console.log(`OKボタンなし（${i}回クリック後）`);
        break;
      }
    }

    await sleep(1500);
    await screenshot(driver, 'debug4b.png');

    // ⑪ 「完了」ボタンをクリック
    console.log('完了ボタンをクリック中...');
    let submitSuccess = false;

    // collect-btnクラスで探す
    try {
      const completeBtn = await driver.findElement(
        By.xpath('//button[normalize-space(text())="完了" and contains(@class,"collect-btn")]')
      );
      await driver.executeScript('arguments[0].scrollIntoView({block:"center"});', completeBtn);
      await sleep(500);
      await driver.executeScript('arguments[0].click();', completeBtn);
      console.log('完了ボタンクリック成功（collect-btn）');
      submitSuccess = true;
    } catch (e) {
      console.log('collect-btnで見つからず...');
    }

    // テキストで探す
    if (!submitSuccess) {
      try {
        const completeBtn = await driver.findElement(
          By.xpath('//button[normalize-space(text())="完了"]')
        );
        await driver.executeScript('arguments[0].scrollIntoView({block:"center"});', completeBtn);
        await sleep(500);
        await driver.executeScript('arguments[0].click();', completeBtn);
        console.log('完了ボタンクリック成功');
        submitSuccess = true;
      } catch (e) {
        console.log('完了ボタンが見つかりません:', e.message);
      }
    }

    await sleep(6000);

    // ⑫ 投稿完了確認
    const finalUrl = await driver.getCurrentUrl();
    console.log('投稿後URL:', finalUrl);
    await screenshot(driver, 'debug_final.png');

    if (
      finalUrl.includes('room.rakuten.co.jp') &&
      !finalUrl.includes('login') &&
      !finalUrl.includes('sign_in') &&
      !finalUrl.includes('collect')
    ) {
      console.log(`✅ 投稿成功: ${item.itemName.substring(0, 40)}`);
      await driver.quit();
      return true;
    } else {
      console.log(`❌ 投稿失敗: ${finalUrl}`);
      await driver.quit();
      return false;
    }

  } catch (e) {
    console.error(`エラー: ${e.message}`);
    await screenshot(driver, 'debug_error.png');
    await driver.quit();
    return false;
  }
}

// ============================================================
// メイン処理
// ============================================================
async function main() {
  console.log('=== 楽天ROOM自動投稿開始（Selenium版）===');

  console.log('EM:', process.env.EM ? '設定済み' : '未設定');
  console.log('PW:', process.env.PW ? '設定済み' : '未設定');
  console.log('PW文字数:', process.env.PW ? process.env.PW.length : 0);
  console.log('SID:', process.env.SID ? '設定済み' : '未設定');
  console.log('GSA:', process.env.GSA ? '設定済み' : '未設定');

  const items = await getUnpostedItems();

  if (items.length === 0) {
    console.log('未投稿の商品がありません');
    return;
  }

  // 今回は1件だけ試す（調査のため）
  const targetItems = items.slice(0, 1);

  for (const item of targetItems) {
    console.log(`\n--- 投稿開始: ${item.itemName.substring(0, 40)} ---`);

    const success = await postToRakutenRoom(item);

    if (success) {
      await markAsPosted(item.rowIndex);
    }

    await sleep(5000);
  }

  console.log('\n=== 楽天ROOM自動投稿完了 ===');
}

main().catch(console.error);
