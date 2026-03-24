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

    for (let attempt = 0; attempt < 5; attempt++) {
      for (const sel of emailSelectors) {
        try {
          await driver.wait(until.elementLocated(sel), 3000);
          emailInput = await driver.findElement(sel);
          if (emailInput) {
            console.log(`メール入力欄発見（${attempt + 1}回目）`);
            break;
          }
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

    await driver.wait(until.elementLocated(By.css('input[type="password"]')), 10000);
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
// OKボタン（<a>タグ）をクリック
// ============================================================
async function clickOkPopup(driver) {
  try {
    const okLink = await driver.findElement(By.css('li.ok a'));
    await driver.executeScript('arguments[0].click();', okLink);
    console.log('OKリンクをクリック');
    await sleep(800);
    return true;
  } catch (e) {
    try {
      const okBtn = await driver.findElement(By.xpath('//*[contains(@ng-click,"ok()") and not(contains(@ng-click,"book"))]'));
      await driver.executeScript('arguments[0].click();', okBtn);
      console.log('OKボタン(ng-click)をクリック');
      await sleep(800);
      return true;
    } catch (e2) {
      return false;
    }
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

    // ② 楽天ROOMのセッション確立
    console.log('楽天ROOMのセッション確立中...');
    await driver.get('https://login.account.rakuten.com/sso/authorize?client_id=rakuten_room_web&redirect_uri=https://room.rakuten.co.jp/common/callback&scope=openid&response_type=code&state=login#/sign_in');
    await sleep(5000);
    console.log('ROOMセッション確立後URL:', await driver.getCurrentUrl());

    // ③ 楽天トップページでセッション確立
    await driver.get('https://www.rakuten.co.jp/?l-id=pc_header_logo');
    await sleep(2000);

    // ④ 商品ページを開く
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

    // ⑤ お気に入りボタンをクリック
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

    // ⑥ お気に入りページへのリダイレクトを待つ
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

    // ⑦ ROOMに投稿リンクを取得
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

    // ⑧ ROOMの投稿フォームにアクセス
    await driver.get(collectUrl);
    await sleep(3000);

    const curUrl = await driver.getCurrentUrl();
    if (curUrl.includes('login') || curUrl.includes('sign_in') || curUrl.includes('403')) {
      console.log('ROOMログインページ/403検出。ログイン中...');
      await performLogin(driver, 'debug_room_relogin', 'room.rakuten.co.jp');
      await sleep(3000);
      await driver.get(collectUrl);
      await sleep(3000);
    }

    console.log('ROOMコレクトフォームURL:', await driver.getCurrentUrl());
    await screenshot(driver, 'debug3.png');

    // ⑨ 投稿フォームに紹介文を入力（AngularJS対応）
    console.log('紹介文を入力中...');
    try {
      await driver.wait(until.elementLocated(By.id('collect-content')), 10000);

      await driver.executeScript(`
        const el = document.querySelector('#collect-content');
        // AngularJSのネイティブvalueセッターを使う
        const nativeInputValueSetter = Object.getOwnPropertyDescriptor(
          window.HTMLTextAreaElement.prototype, 'value'
        ).set;
        nativeInputValueSetter.call(el, arguments[0]);
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('keyup', { bubbles: true }));
        // AngularJSのスコープを強制更新
        try {
          const scope = angular.element(el).scope();
          if (scope) scope.$apply();
        } catch(e) {}
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

    // ⑩ 最初のOKポップアップを閉じる
    console.log('OKポップアップを閉じます（1回目）...');
    for (let i = 0; i < 10; i++) {
      const clicked = await clickOkPopup(driver);
      if (!clicked) {
        console.log(`OKボタンなし（${i}回クリック後）`);
        break;
      }
      console.log(`OKクリック ${i + 1}回目`);
    }

    await sleep(1500);
    await screenshot(driver, 'debug4b.png');

    // ⑪ 「完了」ボタンをクリック
    console.log('完了ボタンをクリック中...');
    let submitSuccess = false;

    try {
      const collectBtns = await driver.findElements(By.css('button.collect-btn'));
      console.log(`collect-btnボタン数: ${collectBtns.length}`);

      for (const btn of collectBtns) {
        const isDisplayed = await btn.isDisplayed();
        const isEnabled = await btn.isEnabled();
        if (isDisplayed && isEnabled) {
          await driver.executeScript('arguments[0].scrollIntoView({block:"center"});', btn);
          await sleep(500);
          await driver.executeScript('arguments[0].click();', btn);
          console.log('✅ collect-btnをクリックしました');
          submitSuccess = true;
          break;
        }
      }
    } catch (e) {
      console.log('collect-btnエラー:', e.message);
    }

    if (!submitSuccess) {
      try {
        const btn = await driver.findElement(By.css('button[ng-click="collect()"]'));
        const isDisplayed = await btn.isDisplayed();
        if (isDisplayed) {
          await driver.executeScript('arguments[0].scrollIntoView({block:"center"});', btn);
          await sleep(500);
          await driver.executeScript('arguments[0].click();', btn);
          console.log('✅ ng-click="collect()"ボタンをクリック');
          submitSuccess = true;
        }
      } catch (e) {
        console.log('ng-click collect()ボタンが見つかりません');
      }
    }

    if (!submitSuccess) {
      console.log('完了ボタンが見つかりませんでした');
      await driver.quit();
      return false;
    }

    // ⑫ 投稿後のポップアップ（成功確認）を閉じる
    console.log('投稿後ポップアップを処理中...');
    await sleep(3000);
    await screenshot(driver, 'debug_after_submit.png');

    // 投稿後のOKポップアップを閉じる（最大10回試みる）
    for (let i = 0; i < 10; i++) {
      const clicked = await clickOkPopup(driver);
      if (clicked) {
        console.log(`投稿後OKポップアップをクリック（${i + 1}回目）`);
        await sleep(1000);
      } else {
        console.log(`投稿後OKポップアップなし（${i}回後）`);
        break;
      }
    }

    // ⑬ 投稿完了を待つ（URLが変わるまで最大15秒）
    console.log('投稿完了を待機中...');
    for (let i = 0; i < 15; i++) {
      await sleep(1000);
      const url = await driver.getCurrentUrl();
      if (url.includes('room.rakuten.co.jp') && !url.includes('collect') && !url.includes('login')) {
        console.log(`✅ 投稿成功（URL変化検知）: ${url}`);
        await screenshot(driver, 'debug_final.png');
        await driver.quit();
        return true;
      }
    }

    // ⑭ 最終確認
    const finalUrl = await driver.getCurrentUrl();
    console.log('投稿後URL:', finalUrl);
    await screenshot(driver, 'debug_final.png');

    // collectページのままでも投稿後ポップアップが出ていれば成功とみなす
    if (finalUrl.includes('room.rakuten.co.jp') && !finalUrl.includes('login') && !finalUrl.includes('sign_in')) {
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

  const targetItems = items.slice(0, 3);

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
