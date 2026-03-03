/**
 * QB_Scrape.js
 *
 * 改善点:
 * ────────────────────────────────────────────
 * 1. 実HTMLのDOM構造に完全対応（セレクタを実際のHTML属性に合わせて修正）
 *    - 解答ボタン: #answerSection / #answerCbtSection 両対応
 *    - 問題文: .body 直下テキスト対応（<p>なしでもOK）
 *    - フッター: 掲載頁 & ID 両方を取得
 *    - 正答率テキストも取得
 * 2. 「医ンプット」セクションの iframe 内テキスト＋画像スクレイピングに対応
 * 3. 「基本事項」「医ンプット」の「すべて表示」ボタンを自動クリック
 * 4. Excel (xlsx) 出力を追加（exceljs使用）
 * 5. JSON 出力を追加（生データ保存、再処理用）
 * 6. 画像をファイルとしても保存（{fileName}_images/ フォルダ）
 * 7. node-fetch 不要（Node.js v22 組み込み fetch 使用）
 * 8. PdfPrinter で Node.js 正規のフォント読み込み（vfs_fonts.js 不要）
 *
 * 実行前の準備:
 *   npm install puppeteer pdfmake exceljs axios image-size
 *   fonts/ フォルダに NotoSansJP-Regular.ttf, NotoSansJP-Bold.ttf を配置
 *   (https://fonts.google.com/noto/specimen/Noto+Sans+JP からDL)
 */

// ── どこから実行しても node_modules を見つけられるようにパスを追加 ──
const os = require('os');
module.paths.unshift(require('path').join(os.homedir(), 'qb-scraper', 'node_modules'));

const puppeteer = require('puppeteer');
const PdfPrinter = require('pdfmake/js/Printer').default;
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { imageSize } = require('image-size');

// ============================================================
// 設定値（ここを環境に合わせて変更）
// ============================================================
const CONFIG = {
  loginUrl: 'https://login.medilink-study.com/login',
  email: 'your-email@example.com',       // ← ログイン用メールアドレス
  password: 'your-password',             // ← ログイン用パスワード
  startUrl: 'https://qb.medilink-study.com/Answer/117A1',  // ← 最初の問題ページURL
  numberOfPages: 100,                    // ← 取得する問題数
  fileName: 'A 消化器',                  // ← 出力フォルダ名 & ファイル名（拡張子不要）
};

// タイムアウト・待機時間
const TIMING = {
  initialWait: 8000,           // ページ遷移後の待機 (ms)
  scrollStep: 100,             // スクロール刻み (px)
  scrollInterval: 80,          // スクロール間隔 (ms)
  retryInterval: 500,          // 問題文取得リトライ間隔 (ms)
  maxRetries: 60,              // 問題文取得の最大リトライ数
  afterClick: 2000,            // ボタンクリック後の待機 (ms)
  selectorTimeout: 10000,      // セレクタ待機のタイムアウト (ms)
};

// PDF レイアウト
const PDF_LAYOUT = {
  availableWidth: 515.28,      // A4 横幅 (pt)
  maxImageCols: 3,             // 画像の最大列数
};

// フォント設定（Node.js PdfPrinter: .ttf パス指定）
// カレントディレクトリ → スクリプトと同じ場所 → ~/qb-scraper の順に探す
const FONT_DIR = [
  path.join(__dirname, 'fonts'),
  path.join(os.homedir(), 'qb-scraper', 'fonts'),
].find(d => fs.existsSync(d)) || path.join(__dirname, 'fonts');
const FONTS = {
  NotoSansJP: {
    normal: path.join(FONT_DIR, 'NotoSansJP-Regular.ttf'),
    bold: path.join(FONT_DIR, 'NotoSansJP-Bold.ttf'),
  }
};

// ============================================================
// ユーティリティ関数
// ============================================================

/** ページ全体をスクロールして lazy-loading 画像を読み込む */
async function autoScroll(page) {
  await page.evaluate(async (step, interval) => {
    await new Promise((resolve) => {
      let totalHeight = 0;
      const timer = setInterval(() => {
        const scrollHeight = document.body.scrollHeight;
        window.scrollBy(0, step);
        totalHeight += step;
        if (totalHeight >= scrollHeight - window.innerHeight) {
          clearInterval(timer);
          resolve();
        }
      }, interval);
    });
  }, TIMING.scrollStep, TIMING.scrollInterval);
}

/** 画像を Base64 data URL + サイズ情報に変換 (pdfmake は PNG/JPEG のみ対応) */
const PDFMAKE_OK_TYPES = new Set(['png', 'jpg']);  // imageSize が返す type 値

async function processImage(src) {
  const FAIL = { dataUrl: null, dimensions: null, buffer: null };
  if (!src) return FAIL;

  if (src.startsWith('data:')) {
    try {
      const base64Data = src.split(',')[1];
      if (!base64Data) return FAIL;
      const buffer = Buffer.from(base64Data, 'base64');
      const dimensions = imageSize(buffer);
      // 実バイナリの形式で判定（MIMEヘッダは信用しない）
      if (!PDFMAKE_OK_TYPES.has(dimensions.type)) {
        console.error(`  processImage: 非対応フォーマット "${dimensions.type}" → スキップ`);
        return FAIL;
      }
      // MIMEを実フォーマットに正規化して返す
      const correctMime = dimensions.type === 'png' ? 'image/png' : 'image/jpeg';
      const dataUrl = `data:${correctMime};base64,${base64Data}`;
      return { dataUrl, dimensions, buffer };
    } catch (error) {
      console.error('  processImage(data) error:', error.message);
      return FAIL;
    }
  }

  try {
    const response = await axios.get(src, { responseType: 'arraybuffer', timeout: 15000 });
    if (response.status === 200) {
      const buffer = Buffer.from(response.data);
      try {
        const dimensions = imageSize(buffer);
        if (!PDFMAKE_OK_TYPES.has(dimensions.type)) {
          console.error(`  processImage: 非対応フォーマット "${dimensions.type}": ${src.substring(0, 80)} → スキップ`);
          return FAIL;
        }
        const mime = dimensions.type === 'png' ? 'image/png' : 'image/jpeg';
        const dataUrl = `data:${mime};base64,${buffer.toString('base64')}`;
        return { dataUrl, dimensions, buffer };
      } catch (sizeErr) {
        console.error(`  processImage: imageSize失敗: ${sizeErr.message} (${src.substring(0, 60)})`);
        return FAIL;
      }
    }
  } catch (error) {
    console.error('  processImage(URL) error:', src.substring(0, 80), error.message);
  }
  return FAIL;
}

/** Puppeteer で画像要素をスクリーンショットして Base64 にする */
async function captureImageElement(page, src) {
  if (!src || !src.startsWith('http')) return src;
  try {
    const escapedSrc = src.replace(/"/g, '\\"');
    const imageElement = await page.$(`img[src="${escapedSrc}"]`);
    if (imageElement) {
      const screenshotData = await imageElement.screenshot({ encoding: 'base64' });
      return `data:image/png;base64,${screenshotData}`;
    }
  } catch (err) {
    console.error('  画像キャプチャエラー:', err.message);
  }
  return src;
}

/** 画像 URL 配列 → すべて Base64 に変換 */
async function captureAllImages(page, imageSrcs) {
  const results = [];
  for (const src of imageSrcs) {
    results.push(await captureImageElement(page, src));
  }
  return results;
}

/** 画像を個別ファイルとして保存 */
async function saveImageToFile(src, filePath) {
  try {
    let buffer;
    if (src.startsWith('data:')) {
      const base64Data = src.split(',')[1];
      buffer = Buffer.from(base64Data, 'base64');
    } else if (src.startsWith('http')) {
      const resp = await axios.get(src, { responseType: 'arraybuffer', timeout: 15000 });
      buffer = Buffer.from(resp.data);
    } else {
      return;
    }
    fs.writeFileSync(filePath, buffer);
  } catch (err) {
    console.error('  画像保存エラー:', filePath, err.message);
  }
}

// ============================================================
// スクレイピング本体
// ============================================================

/**
 * 連続して numPages 件の問題・解説をスクレイピング（4連問除く）
 */
async function scrapeQuestions(page, numPages) {
  const results = [];
  let previousProblemNumber = '';

  for (let i = 0; i < numPages; i++) {
    console.log(`\n━━━ 問題 ${i + 1}/${numPages} ━━━`);

    // ── ページ読み込み待機 + Vue.js レンダリング待ち ──
    // まず question-content か header が出現するまで最大30秒待つ
    try {
      await page.waitForSelector(
        '.question-wrapper, .question-container, .question-content, div.header span',
        { visible: true, timeout: 30000 }
      );
    } catch (e) {
      console.error('  ✗ ページが読み込まれません');
      console.error('  現在のURL:', page.url());
      await page.screenshot({ path: `debug_q${i + 1}.png`, fullPage: true });
      console.error(`  📸 debug_q${i + 1}.png を確認してください`);
      continue;
    }

    // 追加待機 + スクロールで画像のlazy-loadを発火
    await new Promise(r => setTimeout(r, 2000));
    await autoScroll(page);

    // ① 問題番号
    const problemNumber = await page.evaluate(() => {
      const headerEl = document.querySelector('.question-wrapper .header') ||
                       document.querySelector('div.header');
      if (!headerEl) return '';
      // 最初の <span> が問題番号 (例: "102B40")
      const span = headerEl.querySelector('span');
      return span ? span.innerText.trim() : '';
    });
    console.log(`  問題番号: ${problemNumber}`);

    // ── ループ検出: 同じ問題番号なら終了 ──
    if (problemNumber && problemNumber === previousProblemNumber) {
      console.log(`  ⚠ 同じ問題が繰り返されました。スクレイピングを終了します。`);
      break;
    }
    previousProblemNumber = problemNumber;

    // ── 連問（シリアル問題）検出 ──
    const isSerial = await page.evaluate(() => !!document.querySelector('div.pre-body'));

    // ② 問題文 + ③ 問題画像 + ④ 選択肢（連問対応）
    let questionText = '';
    let problemImages = [];
    let choices = [];

    if (isSerial) {
      // ── 連問: pre-body（共有問題文）+ 各小問を一括取得 ──
      serialData = await page.evaluate(() => {
        const preBody = document.querySelector('div.pre-body');
        let sharedText = '';
        const sharedImageSrcs = [];
        if (preBody) {
          const clone = preBody.cloneNode(true);
          clone.querySelectorAll('.images, .figure').forEach(el => el.remove());
          sharedText = clone.innerText.trim();
          preBody.querySelectorAll('.figure img, .images img').forEach(img => {
            const src = (img.getAttribute('src') || img.getAttribute('data-src') || '').trim();
            if (src) sharedImageSrcs.push(src);
          });
        }
        const subQuestions = [];
        document.querySelectorAll('div.question-content').forEach(qc => {
          const serialNum = qc.querySelector('.serialNumber')?.innerText.trim() || '';
          const body = qc.querySelector('.body')?.innerText.trim() || '';
          const opts = Array.from(qc.querySelectorAll('ul.multiple-answer-options li div.ans'))
            .map(el => el.innerText.trim()).filter(Boolean);
          subQuestions.push({ serialNum, body, choices: opts });
        });
        return { sharedText, sharedImageSrcs, subQuestions };
      });
      problemImages = await captureAllImages(page, serialData.sharedImageSrcs);
      const subTexts = serialData.subQuestions.map(sq =>
        `${sq.serialNum}\n${sq.body}`
      ).join('\n\n');
      questionText = serialData.sharedText + '\n\n' + subTexts;
      choices = serialData.subQuestions.flatMap(sq =>
        [`── ${sq.serialNum} ──`, ...sq.choices]
      );
      console.log(`  🔗 連問: ${serialData.subQuestions.length}小問`);
    } else {
      // ── 通常問題 ──
      for (let r = 0; r < TIMING.maxRetries && !questionText; r++) {
        questionText = await page.evaluate(() => {
          const qc = document.querySelector('div.question-content');
          if (!qc) return '';
          const body = qc.querySelector('.body');
          if (!body) return '';
          return body.innerText.trim();
        });
        if (!questionText) {
          await new Promise(r => setTimeout(r, TIMING.retryInterval));
        }
      }
      if (!questionText) {
        console.warn('  ⚠ 問題文が取得できませんでした');
        questionText = '【問題文なし】';
      }
      const problemImageSrcs = await page.evaluate(() => {
        const qc = document.querySelector('div.question-content');
        if (!qc) return [];
        return Array.from(qc.querySelectorAll('div.figure img'))
          .map(img => (img.getAttribute('src') || img.getAttribute('data-src') || '').trim())
          .filter(Boolean);
      });
      problemImages = await captureAllImages(page, problemImageSrcs);
      const choicesRaw = await page.evaluate(() => {
        return Array.from(document.querySelectorAll('ul.multiple-answer-options li div.ans'))
          .map(el => el.innerText.trim())
          .filter(Boolean);
      });
      choices = [...new Set(choicesRaw)];
    }

    // ⑤ フッター情報（掲載頁 & ID）
    const footerData = await page.evaluate(() => {
      const footer = document.querySelector('div.question-footer');
      if (!footer) return { reference: '', problemId: '' };
      const refSpan = footer.querySelector('span');
      const reference = refSpan ? refSpan.innerText.trim() : '';
      const idMatch = footer.innerText.match(/ID\s*[:：]\s*(\d+)/);
      const problemId = idMatch ? idMatch[1] : '';
      return { reference, problemId };
    });

    // ── 問題データまとめ ──
    const problemData = {
      problemNumber,
      questionText,
      problemImages,
      choices,
      reference: footerData.reference,
      problemId: footerData.problemId,
    };

    // ⑥ 「解答を確認する」ボタンをクリック
    //    国試QB: #answerSection / CBT QB: #answerCbtSection 両対応
    try {
      await page.waitForSelector(
        'div#answerSection div.btn, div#answerCbtSection div.btn',
        { visible: true, timeout: 5000 }
      );
      const clicked = await page.evaluate(() => {
        const btn = document.querySelector('div#answerSection div.btn') ||
                    document.querySelector('div#answerCbtSection div.btn');
        if (btn && (btn.textContent.includes('解答を確認する') || btn.textContent.includes('解答'))) {
          btn.click();
          return true;
        }
        // ボタンのテキストが「リトライ」の場合、既に解答済み → そのまま進む
        if (btn) { btn.click(); return true; }
        return false;
      });
      if (!clicked) console.warn('  ⚠ 解答ボタンが見つかりません');
    } catch (error) {
      console.error('  ✗ 解答ボタンエラー:', error.message);
    }

    await new Promise(r => setTimeout(r, TIMING.afterClick));

    // 正解表示を待機
    try {
      await page.waitForSelector('div.resultContent--currentCorrect', {
        visible: true, timeout: TIMING.selectorTimeout
      });
    } catch (e) {
      console.error('  ✗ 正解表示が現れません。スキップします');
      try {
        await page.evaluate(() => {
          document.querySelector('div.toNextWrapper--btn')?.click();
        });
        await new Promise(r => setTimeout(r, TIMING.afterClick));
      } catch (_) { /* skip */ }
      continue;
    }

    // ⑦ 正解 & 正答率（連問: 複数の resultContent ブロック対応）
    const resultData = await page.evaluate(() => {
      const resultBlocks = document.querySelectorAll('div.resultContent');
      if (resultBlocks.length > 1) {
        // 連問: 各小問の正解・正答率を統合
        const subResults = Array.from(resultBlocks).map(block => {
          const num = block.querySelector('.resultContent--number')?.innerText.trim() || '';
          const answer = block.querySelector('span.resultContent--currentCorrectAnswer')?.innerText.trim() || '';
          let rate = '';
          const correctDiv = block.querySelector('.resultContent--currentCorrect');
          if (correctDiv) {
            const spans = correctDiv.querySelectorAll('span');
            for (const span of spans) {
              if (span.innerText.includes('正答率')) {
                rate = span.innerText.replace(/[()（）]/g, '').replace('正答率', '').trim();
                break;
              }
            }
          }
          return { num, answer, rate };
        });
        const correctAnswer = subResults.map(r => `${r.num}: ${r.answer}`).join(' / ');
        const accuracyRate = subResults.map(r => `${r.num}: ${r.rate}`).join(' / ');
        return { correctAnswer, accuracyRate };
      }

      // 通常問題: 単一の正解
      const correctEl = document.querySelector('span.resultContent--currentCorrectAnswer');
      const correctAnswer = correctEl ? correctEl.innerText.trim() : '';
      let accuracyRate = '';
      const correctWrapper = document.querySelector('div.resultContent--currentCorrect');
      if (correctWrapper) {
        const spans = correctWrapper.querySelectorAll('span');
        for (const span of spans) {
          const text = span.innerText.trim();
          if (text.includes('正答率')) {
            accuracyRate = text.replace(/[()（）]/g, '').replace('正答率', '').trim();
            break;
          }
        }
      }
      return { correctAnswer, accuracyRate };
    });

    // ⑧ 解説データの取得
    let explanationData = await page.evaluate(() => {
      // descContent ブロックをタイトルで検索するヘルパー
      const findBlock = (titleText) => {
        return Array.from(document.querySelectorAll('div.descContent')).find(block => {
          return block.querySelector('.descContent--title')?.innerText.trim() === titleText;
        });
      };

      // ヘルパー: ブロック内の全 .descContent--detail を連結取得（連問で複数ある場合に対応）
      const getBlockText = (block) => {
        if (!block) return '';
        const details = block.querySelectorAll('.descContent--detail');
        return Array.from(details).map(d => d.innerText.trim()).filter(Boolean).join('\n\n');
      };

      const pointsBlock = findBlock('解法の要点');
      const explanationPoints = getBlockText(pointsBlock) || '解法の要点なし';

      const optionBlock = findBlock('選択肢解説');
      const optionAnalysis = getBlockText(optionBlock);

      const guidelineBlock = findBlock('ガイドライン');
      const guideline = getBlockText(guidelineBlock);

      // 連問で追加される可能性のあるブロック
      const diagnosisBlock = findBlock('診断');
      const diagnosis = getBlockText(diagnosisBlock);
      const keywordBlock = findBlock('KEYWORD');
      const keyword = getBlockText(keywordBlock);
      const findingsBlock = findBlock('主要所見');
      const findings = getBlockText(findingsBlock);

      // 画像診断ブロック（テキスト + 画像両方取得）
      const imageBlock = findBlock('画像診断');
      const explanationImages = [];
      let imageDiagnosisText = '';
      if (imageBlock) {
        imageDiagnosisText = getBlockText(imageBlock);
        imageBlock.querySelectorAll('img').forEach(img => {
          const src = (img.getAttribute('src') || img.getAttribute('data-src') || '').trim();
          if (src) explanationImages.push(src);
        });
      }

      return { explanationPoints, optionAnalysis, guideline, diagnosis, keyword, findings, imageDiagnosisText, explanationImages };
    });

    // 解説画像をキャプチャ
    if (explanationData.explanationImages?.length > 0) {
      explanationData.explanationImages = await captureAllImages(page, explanationData.explanationImages);
    }

    // ⑨ 基本事項の取得（「すべて表示」を先にクリック → CSS展開 → 画像ロード待ち）
    await page.evaluate(() => {
      // 「すべて表示」ボタンをクリック
      const expandBtn = document.querySelector('div.basic .basic--expandControl');
      if (expandBtn) expandBtn.click();
      // CSS の max-height / overflow 制限を強制解除して完全展開
      const basicMain = document.querySelector('div.basic .basic--main');
      if (basicMain) {
        basicMain.style.maxHeight = 'none';
        basicMain.style.overflow = 'visible';
        basicMain.style.height = 'auto';
      }
      // basicsContent--detail 内も展開
      document.querySelectorAll('div.basic .basicsContent--detail').forEach(el => {
        el.style.maxHeight = 'none';
        el.style.overflow = 'visible';
        el.style.height = 'auto';
      });
      // basic セクション自体も展開
      const basicEl = document.querySelector('div.basic');
      if (basicEl) {
        basicEl.style.maxHeight = 'none';
        basicEl.style.overflow = 'visible';
      }
    });
    await new Promise(r => setTimeout(r, 800));
    // セクションをビューポートにスクロールして lazy-load 画像をトリガー
    await page.evaluate(() => {
      const basicEl = document.querySelector('div.basic');
      if (basicEl) basicEl.scrollIntoView({ behavior: 'instant', block: 'start' });
    });
    await new Promise(r => setTimeout(r, 500));
    // 基本事項内の全画像が読み込まれるまで待つ
    await page.evaluate(() => {
      return Promise.all(
        Array.from(document.querySelectorAll('div.basic img')).map(img => {
          if (img.complete && img.naturalWidth > 0) return Promise.resolve();
          return new Promise(resolve => {
            img.onload = resolve;
            img.onerror = resolve;
            setTimeout(resolve, 5000);
          });
        })
      );
    });

    // 各 basicsContent ブロックの情報を取得（テーブル有無を判定）
    let basicData = await page.evaluate(() => {
      const basicElem = document.querySelector('div.basic');
      if (!basicElem) return null;

      const title = basicElem.querySelector('.basic--title span:first-child')?.innerText.trim() || '';

      // basicsContent ブロックが複数ある可能性（補足事項、基本事項 など）
      const blocks = [];
      basicElem.querySelectorAll('.basicsContent').forEach((block, idx) => {
        const subTitle = block.querySelector('.basicsContent--title')?.innerText.trim() || '';
        const detailEl = block.querySelector('.basicsContent--detail');
        const hasTable = detailEl ? detailEl.querySelector('table') !== null : false;
        const detail = detailEl?.innerText.trim() || '';
        blocks.push({ subTitle, detail, hasTable, blockIndex: idx });
      });

      // img 画像（テーブル以外の通常画像）
      const images = [];
      basicElem.querySelectorAll('.basicsContent--detail img, .basic--main img').forEach(img => {
        const src = (img.getAttribute('src') || '').trim();
        if (src) images.push(src);
      });

      // テキストはテーブルを含まないブロックのみ連結
      const textContent = blocks.map(b => {
        if (b.hasTable) {
          // テーブルブロック: サブタイトルのみテキスト、内容は画像で取得
          return b.subTitle ? `【${b.subTitle}】\n（※ 表は下記画像を参照）` : '';
        }
        return b.subTitle ? `【${b.subTitle}】\n${b.detail}` : b.detail;
      }).filter(Boolean).join('\n\n');

      return { title, textContent, blocks, images };
    });

    // テーブルを含む basicsContent--detail ブロックをスクリーンショット
    if (basicData) {
      const tableScreenshots = [];
      const detailElements = await page.$$('div.basic .basicsContent .basicsContent--detail');
      for (let bi = 0; bi < (basicData.blocks || []).length; bi++) {
        if (basicData.blocks[bi].hasTable && detailElements[bi]) {
          try {
            // テーブル要素をビューポートにスクロール
            await detailElements[bi].evaluate(el => el.scrollIntoView({ behavior: 'instant', block: 'center' }));
            await new Promise(r => setTimeout(r, 300));
            const screenshotData = await detailElements[bi].screenshot({ encoding: 'base64' });
            tableScreenshots.push(`data:image/png;base64,${screenshotData}`);
          } catch (err) {
            console.error('  基本事項テーブル撮影エラー:', err.message);
          }
        }
      }
      // テーブルスクショを images の先頭に追加
      basicData.images = [...tableScreenshots, ...(basicData.images || [])];
    }

    // 基本事項の通常画像をキャプチャ（URLから直接DLして確実に全体を取得）
    if (basicData?.images?.length > 0) {
      const capturedBasicImages = [];
      for (const src of basicData.images) {
        if (src.startsWith('data:')) {
          // すでにBase64（テーブルスクショ等）
          capturedBasicImages.push(src);
        } else if (src.startsWith('http')) {
          try {
            const resp = await axios.get(src, { responseType: 'arraybuffer', timeout: 15000 });
            const buffer = Buffer.from(resp.data);
            const ct = (resp.headers['content-type'] || 'image/png').split(';')[0].trim().toLowerCase();
            if (ct.startsWith('image/') && ct !== 'image/svg+xml') {
              capturedBasicImages.push(`data:${ct};base64,${buffer.toString('base64')}`);
            } else {
              console.error(`  基本事項画像: 非対応Content-Type (${ct})`);
            }
          } catch (err) {
            capturedBasicImages.push(await captureImageElement(page, src));
          }
        } else {
          capturedBasicImages.push(src);
        }
      }
      basicData.images = capturedBasicImages;
    }

    // ⑩ 医ンプットの取得
    let medicalInputData = null;

    const medicalInputVisible = await page.evaluate(() => {
      const titleEl = document.querySelector('div.medical-input .medical-input--title');
      if (!titleEl) return false;
      // display: none なら非表示（コンテンツなし）
      return titleEl.style.display !== 'none';
    });

    if (medicalInputVisible) {
      // 「すべて表示」をクリック + CSS制限解除
      await page.evaluate(() => {
        const expandBtn = document.querySelector('div.medical-input .medical-input--expandControl');
        if (expandBtn) expandBtn.click();
        // CSS の max-height / overflow 制限を強制解除
        const miMain = document.querySelector('div.medical-input .medical-input--main');
        if (miMain) {
          miMain.style.maxHeight = 'none';
          miMain.style.overflow = 'visible';
          miMain.style.height = 'auto';
        }
        const miEl = document.querySelector('div.medical-input');
        if (miEl) {
          miEl.style.maxHeight = 'none';
          miEl.style.overflow = 'visible';
        }
        // iframe自体のサイズ制限も解除
        const iframe = document.querySelector('iframe#medicalInput');
        if (iframe) {
          iframe.style.maxHeight = 'none';
          iframe.style.height = 'auto';
          iframe.style.minHeight = '500px';
        }
      });
      await new Promise(r => setTimeout(r, 1000));
      // セクションをビューポートにスクロール
      await page.evaluate(() => {
        const miEl = document.querySelector('div.medical-input');
        if (miEl) miEl.scrollIntoView({ behavior: 'instant', block: 'start' });
      });
      await new Promise(r => setTimeout(r, 500));

      // iframe 内のコンテンツを取得
      try {
        const iframeEl = await page.$('iframe#medicalInput');
        if (iframeEl) {
          const frame = await iframeEl.contentFrame();
          if (frame) {
            await frame.waitForSelector('body', { timeout: 5000 }).catch(() => {});
            // iframe内の全画像が読み込まれるまで待つ
            await frame.evaluate(() => {
              return Promise.all(
                Array.from(document.querySelectorAll('img')).map(img => {
                  if (img.complete && img.naturalWidth > 0) return Promise.resolve();
                  return new Promise(resolve => {
                    img.onload = resolve;
                    img.onerror = resolve;
                    setTimeout(resolve, 5000);
                  });
                })
              );
            }).catch(() => {});
            medicalInputData = await frame.evaluate(() => {
              const text = document.body?.innerText?.trim() || '';
              const images = [];
              document.querySelectorAll('img').forEach(img => {
                let src = img.getAttribute('src') || '';
                // 相対URLを絶対URLに変換
                if (src && !src.startsWith('http') && !src.startsWith('data:')) {
                  try {
                    src = new URL(src, document.baseURI).href;
                  } catch (_) { /* keep as is */ }
                }
                if (src) images.push(src.trim());
              });
              return { text, images };
            });

            // 医ンプット画像をキャプチャ（iframe内の画像は frameでスクショ）
            if (medicalInputData?.images?.length > 0) {
              const capturedImages = [];
              for (const src of medicalInputData.images) {
                if (src.startsWith('http')) {
                  // iframe内からはpage.$が使えないので、axiosで直接取得
                  try {
                    const resp = await axios.get(src, { responseType: 'arraybuffer', timeout: 15000 });
                    const buffer = Buffer.from(resp.data);
                    const ct = (resp.headers['content-type'] || 'image/png').split(';')[0].trim().toLowerCase();
                    if (ct.startsWith('image/') && ct !== 'image/svg+xml') {
                      capturedImages.push(`data:${ct};base64,${buffer.toString('base64')}`);
                    } else {
                      console.error(`  医ンプット画像: 非対応Content-Type (${ct})`);
                    }
                  } catch (err) {
                    capturedImages.push(src);
                  }
                } else {
                  capturedImages.push(src);
                }
              }
              medicalInputData.images = capturedImages;
            }

            // テキストが空 or 極端に短い場合はnullにする
            if (medicalInputData && !medicalInputData.text && medicalInputData.images.length === 0) {
              medicalInputData = null;
            }
          }
        }
      } catch (err) {
        console.error('  ⚠ 医ンプット取得エラー:', err.message);
      }
    }

    // ── 結果をまとめる ──
    const combinedData = {
      problem: problemData,
      isSerial: isSerial || false,
      subQuestions: isSerial ? (serialData ? serialData.subQuestions : []) : [],
      result: resultData,
      explanation: explanationData,
      basic: basicData,
      medicalInput: medicalInputData,
    };

    console.log(`  ✓ 完了 (正解: ${resultData.correctAnswer})`);
    results.push(combinedData);

    // ⑪ 「次の問題へ」ボタンをクリック
    if (i < numPages - 1) {
      try {
        await page.waitForSelector('div.toNextWrapper--btn', {
          visible: true, timeout: TIMING.selectorTimeout
        });
        await page.evaluate(() => {
          document.querySelector('div.toNextWrapper--btn')?.click();
        });
        await new Promise(r => setTimeout(r, TIMING.afterClick));
        await page.waitForSelector('.question-wrapper .header, div.header', {
          visible: true, timeout: TIMING.selectorTimeout
        });
      } catch (err) {
        console.error('  ✗ 次の問題への遷移エラー:', err.message);
        break;
      }
    }
  }

  return results;
}

// ============================================================
// 画像ファイル保存
// ============================================================

/**
 * 全問題の画像をファイルとして保存する
 * 戻り値: { problemNumber -> { type -> [filePath, ...] } }
 */
async function saveAllImages(results, outputDir) {
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const imageMap = {};

  for (const data of results) {
    const pn = data.problem.problemNumber || 'unknown';
    imageMap[pn] = {};

    // 問題画像
    if (data.problem.problemImages?.length > 0) {
      imageMap[pn].problem = [];
      for (let j = 0; j < data.problem.problemImages.length; j++) {
        const fName = `${pn}_問題_${j + 1}.png`;
        const fPath = path.join(outputDir, fName);
        await saveImageToFile(data.problem.problemImages[j], fPath);
        imageMap[pn].problem.push(fName);
      }
    }

    // 解説画像
    if (data.explanation?.explanationImages?.length > 0) {
      imageMap[pn].explanation = [];
      for (let j = 0; j < data.explanation.explanationImages.length; j++) {
        const fName = `${pn}_解説_${j + 1}.png`;
        const fPath = path.join(outputDir, fName);
        await saveImageToFile(data.explanation.explanationImages[j], fPath);
        imageMap[pn].explanation.push(fName);
      }
    }

    // 基本事項画像
    if (data.basic?.images?.length > 0) {
      imageMap[pn].basic = [];
      for (let j = 0; j < data.basic.images.length; j++) {
        const fName = `${pn}_基本事項_${j + 1}.png`;
        const fPath = path.join(outputDir, fName);
        await saveImageToFile(data.basic.images[j], fPath);
        imageMap[pn].basic.push(fName);
      }
    }

    // 医ンプット画像
    if (data.medicalInput?.images?.length > 0) {
      imageMap[pn].medicalInput = [];
      for (let j = 0; j < data.medicalInput.images.length; j++) {
        const fName = `${pn}_医ンプット_${j + 1}.png`;
        const fPath = path.join(outputDir, fName);
        await saveImageToFile(data.medicalInput.images[j], fPath);
        imageMap[pn].medicalInput.push(fName);
      }
    }
  }

  return imageMap;
}

// ============================================================
// PDF 生成
// ============================================================

/** 画像配列 → pdfmake content 要素 (共通ヘルパー) */
async function buildImageContent(imageSrcs, errorLabel = '画像', maxScale = 0.5) {
  const fullWidth = PDF_LAYOUT.availableWidth;
  const maxW = fullWidth * maxScale;  // 画像の最大幅を availableWidth の割合で制限
  const getW = (obj) => {
    if (!obj?.dimensions?.width) return maxW;
    // 元の幅をpt換算（スクショは96dpiなので px * 0.75 ≈ pt）
    const origPt = obj.dimensions.width * 0.75;
    return Math.min(origPt, maxW);
  };

  if (!imageSrcs || imageSrcs.length === 0) return [];

  const processed = [];
  for (const src of imageSrcs) {
    processed.push(await processImage(src));
  }

  const total = processed.length;

  if (total === 1) {
    return processed[0].dataUrl
      ? [{ image: processed[0].dataUrl, width: getW(processed[0]), margin: [0, 5, 0, 5] }]
      : [{ text: `${errorLabel}読み込みエラー`, style: 'error' }];
  }

  // 2枚以上: 上段 ceil(n/2) 枚, 下段 残り
  const firstRowCount = Math.ceil(total / 2);
  const makeCell = (p) => p?.dataUrl
    ? { image: p.dataUrl, width: getW(p) }
    : { text: `${errorLabel}読み込みエラー`, style: 'error' };

  const row1 = processed.slice(0, firstRowCount).map(makeCell);
  const tableBody = [row1];

  if (total > firstRowCount) {
    const row2 = processed.slice(firstRowCount).map(makeCell);
    while (row2.length < firstRowCount) row2.push({ text: '' });
    tableBody.push(row2);
  }

  return [{
    table: { widths: row1.map(() => '*'), body: tableBody },
    layout: 'noBorders',
    margin: [0, 5, 0, 5],
  }];
}

/** PDF 生成 */
async function generatePdf(results, fileName) {
  if (!fs.existsSync(FONTS.NotoSansJP.normal)) {
    console.error(`\n❌ フォントなし: ${FONTS.NotoSansJP.normal}`);
    console.error('   fonts/ に NotoSansJP-Regular.ttf, NotoSansJP-Bold.ttf を配置してください\n');
    return;
  }

  const printer = new PdfPrinter(FONTS);
  const doc = {
    content: [],
    defaultStyle: { font: 'NotoSansJP', fontSize: 10.5 },
    styles: {
      h1:            { fontSize: 14, bold: true, margin: [0, 0, 0, 8] },
      h2:            { fontSize: 12, bold: true, margin: [0, 12, 0, 5] },
      body:          { fontSize: 10.5, margin: [0, 3, 0, 3] },
      choices:       { fontSize: 10.5, margin: [10, 2, 0, 2] },
      correctLarge:  { fontSize: 13, bold: true, color: '#D32F2F', margin: [0, 5, 0, 5] },
      detail:        { fontSize: 10.5, margin: [10, 0, 0, 5] },
      error:         { fontSize: 10.5, color: 'red', margin: [0, 5, 0, 5] },
      small:         { fontSize: 9, color: '#666', margin: [0, 2, 0, 2] },
    },
  };

  for (const data of results) {
    const pn = data.problem.problemNumber || '???';

   try {
    // ── 問題ページ ──
    doc.content.push(
      { text: pn, style: 'h1' },
      { text: data.problem.reference ? `[${data.problem.reference}]` : '', style: 'small' },
      { text: data.problem.questionText, style: 'body' }
    );

    // 問題画像
    if (data.problem.problemImages?.length > 0) {
      doc.content.push(...await buildImageContent(data.problem.problemImages, '問題画像', 0.55));
    }

    // 選択肢 (pdfmakeが配列を変更する可能性があるためコピーを渡す)
    if (data.problem.choices?.length > 0) {
      doc.content.push({ ul: [...data.problem.choices], style: 'choices' });
    }

    doc.content.push({ text: '', pageBreak: 'after' });

    // ── 解説ページ ──
    // 正解
    const answerLine = `正解: ${data.result.correctAnswer}` +
      (data.result.accuracyRate ? `  (正答率 ${data.result.accuracyRate})` : '');
    doc.content.push(
      { text: `${pn}  解説`, style: 'h1' },
      { text: answerLine, style: 'correctLarge' }
    );

    // 主要所見
    if (data.explanation?.findings) {
      doc.content.push(
        { text: '主要所見', style: 'h2' },
        { text: data.explanation.findings, style: 'detail' }
      );
    }

    // KEYWORD
    if (data.explanation?.keyword) {
      doc.content.push(
        { text: 'KEYWORD', style: 'h2' },
        { text: data.explanation.keyword, style: 'detail' }
      );
    }

    // 画像診断（テキスト + 画像）
    if (data.explanation?.imageDiagnosisText || data.explanation?.explanationImages?.length > 0) {
      doc.content.push({ text: '画像診断', style: 'h2' });
      if (data.explanation?.imageDiagnosisText) {
        doc.content.push({ text: data.explanation.imageDiagnosisText, style: 'detail' });
      }
      if (data.explanation?.explanationImages?.length > 0) {
        doc.content.push(...await buildImageContent(data.explanation.explanationImages, '解説画像', 0.5));
      }
    }

    // 診断
    if (data.explanation?.diagnosis) {
      doc.content.push(
        { text: '診断', style: 'h2' },
        { text: data.explanation.diagnosis, style: 'detail' }
      );
    }

    // 解法の要点
    if (data.explanation?.explanationPoints) {
      doc.content.push(
        { text: '解法の要点', style: 'h2' },
        { text: data.explanation.explanationPoints, style: 'detail' }
      );
    }

    // 選択肢解説
    if (data.explanation?.optionAnalysis) {
      doc.content.push(
        { text: '選択肢解説', style: 'h2' },
        { text: data.explanation.optionAnalysis, style: 'detail' }
      );
    }

    // ガイドライン
    if (data.explanation?.guideline) {
      doc.content.push(
        { text: 'ガイドライン', style: 'h2' },
        { text: data.explanation.guideline, style: 'detail' }
      );
    }

    // 基本事項
    if (data.basic) {
      doc.content.push(
        { text: data.basic.title || '基本事項など', style: 'h2' },
        { text: data.basic.textContent, style: 'detail' }
      );
      if (data.basic.images?.length > 0) {
        doc.content.push(...await buildImageContent(data.basic.images, '基本事項画像', 0.65));
      }
    }

    // 医ンプット
    if (data.medicalInput) {
      doc.content.push({ text: '医ンプット', style: 'h2' });
      if (data.medicalInput.text) {
        doc.content.push({ text: data.medicalInput.text, style: 'detail' });
      }
      if (data.medicalInput.images?.length > 0) {
        doc.content.push(...await buildImageContent(data.medicalInput.images, '医ンプット画像', 0.5));
      }
    }

    doc.content.push({ text: '', pageBreak: 'after' });
   } catch (contentErr) {
    console.error(`  ⚠ PDF内容構築エラー (${pn}):`, contentErr.message);
    // エラーが出た問題はスキップし、区切りだけ入れて続行
    doc.content.push(
      { text: `${pn}: PDF構築エラー (${contentErr.message})`, style: 'error' },
      { text: '', pageBreak: 'after' }
    );
   }
  }

  // 書き出し
  const outputPath = path.join(fileName, `${fileName}.pdf`);
  const pdfDoc = await printer.createPdfKitDocument(doc);
  return new Promise((resolve, reject) => {
    try {
      const stream = fs.createWriteStream(outputPath);
      pdfDoc.pipe(stream);
      pdfDoc.end();
      stream.on('finish', () => { console.log(`  ✅ PDF: ${outputPath}`); resolve(outputPath); });
      stream.on('error', reject);
    } catch (err) {
      console.error('  PDF生成エラー:', err);
      reject(err);
    }
  });
}

// ============================================================
// Excel 生成
// ============================================================

async function generateExcel(results, fileName) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'QB Scraper Ver.5';

  // ── シート: 問題一覧 ──
  const ws = workbook.addWorksheet('問題一覧');

  // ヘッダー行
  const headers = [
    '#', '問題番号', '掲載頁', '問題文', '選択肢',
    '正解', '正答率',
    '解法の要点', '選択肢解説', 'ガイドライン',
    '基本事項', '医ンプット',
    '問題画像ファイル', '解説画像ファイル',
  ];
  const headerRow = ws.addRow(headers);
  headerRow.font = { bold: true, size: 11 };
  headerRow.fill = {
    type: 'pattern', pattern: 'solid',
    fgColor: { argb: 'FF4472C4' },
  };
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };

  // 列幅設定
  ws.columns = [
    { width: 5 },   // #
    { width: 12 },  // 問題番号
    { width: 16 },  // 掲載頁
    { width: 50 },  // 問題文
    { width: 40 },  // 選択肢
    { width: 8 },   // 正解
    { width: 14 },  // 正答率
    { width: 60 },  // 解法の要点
    { width: 60 },  // 選択肢解説
    { width: 30 },  // ガイドライン
    { width: 50 },  // 基本事項
    { width: 50 },  // 医ンプット
    { width: 25 },  // 問題画像
    { width: 25 },  // 解説画像
  ];

  // データ行
  results.forEach((data, idx) => {
    const pn = data.problem.problemNumber || '';
    const imgDir = path.join(fileName, `${fileName}_images`);
    const probImgFiles = data.problem.problemImages?.length > 0
      ? data.problem.problemImages.map((_, j) => `${imgDir}/${pn}_問題_${j + 1}.png`).join(', ')
      : '';
    const explImgFiles = data.explanation?.explanationImages?.length > 0
      ? data.explanation.explanationImages.map((_, j) => `${imgDir}/${pn}_解説_${j + 1}.png`).join(', ')
      : '';

    const row = ws.addRow([
      idx + 1,
      pn,
      data.problem.reference || '',
      data.problem.questionText || '',
      data.problem.choices?.join('\n') || '',
      data.result?.correctAnswer || '',
      data.result?.accuracyRate || '',
      data.explanation?.explanationPoints || '',
      data.explanation?.optionAnalysis || '',
      data.explanation?.guideline || '',
      data.basic?.textContent || '',
      data.medicalInput?.text || '',
      probImgFiles,
      explImgFiles,
    ]);

    // テキスト折り返し
    row.alignment = { wrapText: true, vertical: 'top' };
  });

  // オートフィルター
  ws.autoFilter = { from: 'A1', to: `N${results.length + 1}` };

  // ウィンドウ枠固定（ヘッダー行）
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  // 保存
  const outputPath = path.join(fileName, `${fileName}.xlsx`);
  await workbook.xlsx.writeFile(outputPath);
  console.log(`  ✅ Excel: ${outputPath}`);
}

// ============================================================
// JSON 保存
// ============================================================

/** 画像参照を要約文字列に変換するヘルパー */
function summarizeImageSrc(s) {
  if (typeof s !== 'string') return '[非テキスト]';
  return s.startsWith('data:') ? `[Base64 ${Math.round(s.length / 1024)}KB]` : s;
}

function saveJson(results, fileName) {
  // Base64画像データはJSONが巨大になるため、URLのみ保持するバージョンを保存
  // pdfmakeがresults内のオブジェクトを変更する場合があるため、明示的プロパティ参照を使用
  const lite = results.map(data => ({
    problem: {
      problemNumber: data.problem.problemNumber || '',
      questionText: data.problem.questionText || '',
      problemImages: (data.problem.problemImages || []).map(summarizeImageSrc),
      choices: (data.problem.choices || []).map(c => typeof c === 'string' ? c : String(c)),
      reference: data.problem.reference || '',
      problemId: data.problem.problemId || '',
    },
    isSerial: data.isSerial || false,
    subQuestions: Array.isArray(data.subQuestions) ? data.subQuestions.map(sq => ({
      serialNum: sq.serialNum || '',
      body: sq.body || '',
      choices: Array.isArray(sq.choices) ? sq.choices.map(c => typeof c === 'string' ? c : String(c)) : [],
    })) : [],
    result: {
      correctAnswer: data.result?.correctAnswer || '',
      accuracyRate: data.result?.accuracyRate || '',
    },
    explanation: data.explanation ? {
      explanationPoints: data.explanation.explanationPoints || '',
      optionAnalysis: data.explanation.optionAnalysis || '',
      guideline: data.explanation.guideline || '',
      diagnosis: data.explanation.diagnosis || '',
      keyword: data.explanation.keyword || '',
      findings: data.explanation.findings || '',
      imageDiagnosisText: data.explanation.imageDiagnosisText || '',
      explanationImages: (data.explanation.explanationImages || []).map(summarizeImageSrc),
    } : null,
    basic: data.basic ? {
      title: data.basic.title || '',
      textContent: data.basic.textContent || '',
      images: (data.basic.images || []).map(summarizeImageSrc),
    } : null,
    medicalInput: data.medicalInput ? {
      text: data.medicalInput.text || '',
      images: (data.medicalInput.images || []).map(summarizeImageSrc),
    } : null,
  }));

  const outputPath = path.join(fileName, `${fileName}.json`);
  fs.writeFileSync(outputPath, JSON.stringify(lite, null, 2), 'utf-8');
  console.log(`  ✅ JSON: ${outputPath}`);
}

// ============================================================
// メイン
// ============================================================

async function main() {
  console.log('🚀 QB スクレイピング Ver.5');
  console.log(`   URL   : ${CONFIG.startUrl}`);
  console.log(`   問題数: ${CONFIG.numberOfPages}`);
  console.log(`   出力名: ${CONFIG.fileName}`);

  const browser = await puppeteer.launch({
    headless: true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--lang=ja-JP',
      '--font-render-hinting=none',
    ],
  });
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 900 });

  try {
    // ── ログイン ──
    console.log('\n📝 ログイン中...');
    await page.goto(CONFIG.loginUrl, { waitUntil: 'networkidle2' });
    await page.type('input[name="username"]', CONFIG.email);
    await page.type('input[name="password"]', CONFIG.password);
    await page.click('button[type="submit"]');
    await page.waitForNavigation({ waitUntil: 'networkidle2' });
    console.log('  ✓ ログイン完了');

    // ── 最初の問題ページへ ──
    await page.goto(CONFIG.startUrl, { waitUntil: 'networkidle2', timeout: 30000 });

    // Vue.js SPAのレンダリング完了を待つ
    console.log('  ⏳ ページ描画待ち...');
    try {
      await page.waitForSelector('.question-wrapper, .question-container, div.header', {
        visible: true, timeout: 30000
      });
      console.log('  ✓ ページ描画完了');
    } catch (e) {
      console.error('  ⚠ ページ描画タイムアウト。スクリーンショットを保存します...');
      await page.screenshot({ path: 'debug_initial_page.png', fullPage: true });
      console.log('  📸 debug_initial_page.png を確認してください');
      console.log('  現在のURL:', page.url());
    }

    // ── スクレイピング ──
    const results = await scrapeQuestions(page, CONFIG.numberOfPages);
    console.log(`\n📊 スクレイピング完了: ${results.length} 問`);

    // ── 画像ファイル保存 ──
    const outDir = CONFIG.fileName;
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
    const imageDir = path.join(outDir, `${CONFIG.fileName}_images`);
    console.log(`\n🖼️  画像保存中... → ${imageDir}/`);
    await saveAllImages(results, imageDir);

    // ── 各フォーマットで出力（個別エラーハンドリング）──
    console.log('\n📄 出力ファイル生成中...');
    try { await generatePdf(results, CONFIG.fileName); } catch (e) { console.error('  ❌ PDF生成エラー:', e.message); }
    try { await generateExcel(results, CONFIG.fileName); } catch (e) { console.error('  ❌ Excel生成エラー:', e.message); }
    try { saveJson(results, CONFIG.fileName); } catch (e) { console.error('  ❌ JSON保存エラー:', e.message); }

  } catch (error) {
    console.error('❌ エラー:', error);
  } finally {
    await browser.close();
    console.log('\n🏁 完了');
  }
}

main();
