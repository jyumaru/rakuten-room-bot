const puppeteer = require('puppeteer');
const { google } = require('googleapis');

// ============================================================
// 設定（コマンドライン引数から取得）
// node post.js <email> <pass> <sheetId> <gsa_base64>
// ============================================================
const args = process.argv.slice(2);
const CFG = {
  A: args[0], // jyumaru.shidou@gmail.com
  B: args[1], // e4KwbXGJH7aR
  C: args[2], // 1iDIrzBRZQt6SUYtSI1Cro8YyWq8SPcq
  D: args[3], // ewogICJ0eXBlIjogInNlcnZpY2VfYWNjb3VudCIsCiAgInByb2plY3RfaWQiOiAicmFrdXRlbnJvb21ib3QtNDkxMTE3IiwKICAicHJpdmF0ZV9rZXlfaWQiOiAiYjdlODljNTlkZmQwMDk2MDViZTIxYjVhZWNmNDY4NGZhM2VkNDUyYSIsCiAgInByaXZhdGVfa2V5IjogIi0tLS0tQkVHSU4gUFJJVkFURSBLRVktLS0tLVxuTUlJRXZnSUJBREFOQmdrcWhraUc5dzBCQVFFRkFBU0NCS2d3Z2dTa0FnRUFBb0lCQVFEQURrY3hJZ3c5TXIxclxucXFPTmlKQytWbnI4UUw1YzZncGhaUUc5bUhEMDlEV29RdHRFVC9qVjYzd3lwK29RMlhrU0IwMUlKZ3l4NlhMN1xuNFJCNS9WUzVXenQwYjhmQm5xZXVyWU9sNENsWGJlZ1Yvb1VlNDFvTWdIMk9odEk2Y25uQ2NXTEd4cFVGYWc5N1xucGUzSnhuT3hwTkF2cHNyTm9UUDhFSVB0enl1ODBBMnVGazB6U29DOWQ0WnpHMUJydHRlR3Y5NGVqcE5PWEpiL1xuMk9XMVFLZ0hFQ0dJRXFqcGVkT3JKRjNkSmkzb3I5K0JSSUdPQ3VWRjBDaXc4WGVocGJFS0lvemUzR0E1U3VsNVxuU0x6U2NrSlR3M09kZXNKWGdzRFpkSWc3RFhndldUc0RWRUpyTTFEcEhyUDU2WFNNU3kzVytvOEFSSmN6SlhWWVxuMkp0Vm1iVFpBZ01CQUFFQ2dnRUFDK0cyUU5GbzdXNVNyck1IWEUyN0dyTng5Mkl3Qk1LTDh0dVRZSDNxV3hVeVxucnA2NFB3RXRsVnAwdkJPTVZRK0hRSGpJTExNQjdRM1N5Y2R5UkFIS3VJN3UzalMrS0huZXlOMCtQRWhpZG1DN1xuejRTSUN0R285QVlNL29EVHg3N21Ub1BDUnlibytFVklBTy9TaHIycTBhNHZIUXVXRjJPbU9pMXhaQWlCbmJRclxuNDF2MlAyUC9iaDJWdDlvZXQ0akFjcHhNbHhqWHVCV0RVcGd4QWkvM0kycjBEcENtelBYQmx4OFNMYTMva1NVOFxuUjFTdnJCOWhYazU3amhnYTJNQ3dnaTBlbytBeVJBVGsyRTNoL1BDVjFVemV6Z25aQXpERW5CdmlpK0pzZTVqUVxueW1LWVlxY3dIbFBwQlJpeDFqY3djU3MxL2RYY1MrWXdQM1Y4alQvSTZ3S0JnUUR4ZE5mQVQ3NW5tb3lESGovY1xuWlIwektsaGdjblBzNTB0Umt1SGxrK2g4cWVuQ0hmU1hBMUU1S0hVQWVTTmFzU3RMWW93Y3R5WFE2L1hrZ2M5QVxuRWFXOTBuZmppQnVPNEhUTHNwWXlZUzV1L1VQZXlqUE5pSmhmUUpMbytBTHRKMXZxc0kxZ3NzdmtDVE5OcTNuc1xuLzVhdlJ0aUdhMEhnODZ5dmQvbTl2ak1YNndLQmdRRExuN0t3YWcwS2N3S0M4TnFvM1M2M1Q4Sy9KNXdSQk1hZFxuUnh5WC9TckFXWE5QWURrU0ZOaktzWjZFNW5wbGFHYUhaS1pWODFZQlpiTmhpZ3QxZXVDVnJ6YXZkclpUMy9JQVxuUndPNFV4eHcwTW1DNWoyei9qQ0NzSzl5b3Z2ZkdiYUNlYVdHNE04TlQ4MGlQY2pvY1hWb2FQSHhSOVRsVFYwTFxubTE3WllMTlpTd0tCZ1FDVnFMTjF5cmVjNWNsRUdBTERLNVV1dW9kdXVHSXNLNnlla2lrY01GSkF1dHhkNmsxSlxuTU5BdVdtb3k4ZUs4K3VWMzQwd3ZIRUgvUGRINllZOUJDZTh1T2Y3L2M4U0pDWXk3R1NWSmNyemlKRzdsNzNTdVxuWjRUeVBVY1J5VytlNk85ckJ5V0tFeWlYWGpDRGFzNjIzRERjMFUreCtWY3JCRDQ3d0dSMmZDYVZJd0tCZ1FDVFxuWUxEcWNyZWhtb0IwMlhMSnlkem9ITGl0dGpPRk5kbXpPQ2IvOHVNZ2VSMjJrOFI2eTgvbFZRMlF6MmhEUVg4RFxuKzl0UVZtRW5mYjZKbUdxV3l5c0Y2OTArdmtOVkRiK1FaOVhQY1lnaU4xdkNmSGFvY2hBV1oxOTFMM1h4a2lEQVxuNnQ3ZGNwVXA0MXByc0NCYjdOSzNrVTJiL3d1ZU01Sm10anUrUmZsSlpRS0JnQ2QxaFBubjNsNkVYRTFUVVF0RlxueUhnVEwvOUJ3cE1TNXY0Q2VMQktnWjFGK05MS3hzUFZRWkMxTlVCTkNURlRiVnZMNDBIYjZsT3lUR1BCU2Q5dFxuTlc2NDc1NVI5WW1WbXRTWUhoczZqUGxnVW1XNzVvQzVpcU8rc3MrSVlCTzdMc3RwcHB6dWRlUG11NTg5SmRCQlxuVEc2ZVRZNW80dUxIVHAxaFVkTFo4ckNmXG4tLS0tLUVORCBQUklWQVRFIEtFWS0tLS0tXG4iLAogICJjbGllbnRfZW1haWwiOiAicmFrdXRlbi1yb29tLWJvdEByYWt1dGVucm9vbWJvdC00OTExMTcuaWFtLmdzZXJ2aWNlYWNjb3VudC5jb20iLAogICJjbGllbnRfaWQiOiAiMTAyNTY3MTU5ODUwOTYwNzY2OTEyIiwKICAiYXV0aF91cmkiOiAiaHR0cHM6Ly9hY2NvdW50cy5nb29nbGUuY29tL28vb2F1dGgyL2F1dGgiLAogICJ0b2tlbl91cmkiOiAiaHR0cHM6Ly9vYXV0aDIuZ29vZ2xlYXBpcy5jb20vdG9rZW4iLAogICJhdXRoX3Byb3ZpZGVyX3g1MDlfY2VydF91cmwiOiAiaHR0cHM6Ly93d3cuZ29vZ2xlYXBpcy5jb20vb2F1dGgyL3YxL2NlcnRzIiwKICAiY2xpZW50X3g1MDlfY2VydF91cmwiOiAiaHR0cHM6Ly93d3cuZ29vZ2xlYXBpcy5jb20vcm9ib3QvdjEvbWV0YWRhdGEveDUwOS9yYWt1dGVuLXJvb20tYm90JTQwcmFrdXRlbnJvb21ib3QtNDkxMTE3LmlhbS5nc2VydmljZWFjY291bnQuY29tIiwKICAidW5pdmVyc2VfZG9tYWluIjogImdvb2dsZWFwaXMuY29tIgp9Cg==
};

// ============================================================
// Googleスプレッドシートから未投稿の商品を取得
// ============================================================
async function getUnpostedItems() {
  const json = Buffer.from(CFG.D, 'base64').toString('utf8');
  const credentials = JSON.parse(json);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: CFG.C,
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
  const json = Buffer.from(CFG.D, 'base64').toString('utf8');
  const credentials = JSON.parse(json);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.update({
    spreadsheetId: CFG.C,
    range: `楽天ROOM投稿リスト!H${rowIndex}`,
    valueInputOption: 'RAW',
    requestBody: { values: [['済']] },
  });

  console.log(`行${rowIndex}を「済」に更新しました`);
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
      await emailInput.type(CFG.A);

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
        await passInput.type(CFG.B);
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
        if (emailInput) await emailInput.type(CFG.A);

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
          await passInput.type(CFG.B);
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

  if (!CFG.A || !CFG.B || !CFG.C || !CFG.D) {
    console.error('引数が不足しています');
    process.exit(1);
  }

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
