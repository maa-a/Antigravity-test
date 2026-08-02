/**
 * AI会計コンバーター V8.00
 * （文字化け修正 ＆ PCA会計仕訳窓 ＆ Excel／スプレッドシート出力 対応版）
 * ------------------------------------------------------------------
 * 【V7.00 からの変更点】
 * 1. 文字化け対策
 *    ・CSVのクォート処理を追加（摘要に「,」「"」「改行」が入ると列がズレて
 *      文字化けに見えていた根本原因を修正）。
 *    ・Geminiの応答が複数パートに分割されたとき先頭しか読んでおらず
 *      JSONが途中で切れていた不具合を修正（全パートを連結）。
 *    ・応答形式を application/json に固定し、Markdown枠混入を防止。
 *    ・制御文字・置換文字（U+FFFD）・Excel由来の "_x000D_" を除去。
 * 2. WEB画面上の提供元クレジット表記を全削除（index.html 側）。
 * 3. 窓5「PCA会計 仕訳変換窓」を新設。
 *    ・複数ファイルをまとめて1つのCSVに出力。
 *    ・社内用（社内財務管理表 取込用）／会計事務所用（PCA仕訳インポート用）
 *      の2系統を出し分け。
 * 4. Excel(.xlsx)／Googleスプレッドシート出力に対応。
 * 5. 科目が不明な場合は【必ず空欄】で出力（推測や化けた文字列を出さない）。
 *
 * 【設計方針（V7.00 から継承）】
 * 窓1(CashWindow)・窓2(SmartCreditWindow)・窓3(CreditWindow) は
 * それぞれ独立クラスのままで、処理ロジック・プロンプト文言・出力
 * フォーマットは V7.00 と互換です（CSVのクォート処理のみ追加）。
 * ------------------------------------------------------------------
 */

// ====================================================================
// エントリーポイント（index.html の google.script.run から呼ばれる関数）
// GASの仕様上トップレベル関数である必要があるため、ここでは各クラスへの
// 薄い橋渡しだけを行い、実処理はすべてクラス側に委譲しています。
// ====================================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('AI Accounting Converter V8.00')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setUserApiKey(key) {
  PropertiesService.getScriptProperties().setProperty('USER_GEMINI_API_KEY', key);
}

function getUserApiKey() {
  return PropertiesService.getScriptProperties().getProperty('USER_GEMINI_API_KEY') || '';
}

/** 🏦 窓1：通帳・現金（スマート取引取込用） */
function processCash(base64Data, mimeType, customPrompt, modelName) {
  return new CashWindow().process(base64Data, mimeType, customPrompt, modelName);
}

/** 💳 窓2：クレジット（スマート取引取込用・使用目的推測付き） */
function processSmartCredit(base64Data, mimeType, customPrompt, modelName) {
  return new SmartCreditWindow().process(base64Data, mimeType, customPrompt, modelName);
}

/** 💳 窓3：クレジット（PC版弥生インポート用 中間ファイル生成） */
function processCredit(base64Data, mimeType, customPrompt, modelName) {
  return new CreditWindow().process(base64Data, mimeType, customPrompt, modelName);
}

/** 🧾 窓5：請求書・領収書 → PCA会計 仕訳（1ファイル分の解析） */
function processInvoiceForPca(base64Data, mimeType, customPrompt, modelName, sourceFileName) {
  return new PcaInvoiceWindow().process(base64Data, mimeType, customPrompt, modelName, sourceFileName);
}

/**
 * 🧾 窓5：蓄積した全明細から最終CSVを組み立てる（複数ファイル→1CSV）
 * @param voucherStart 伝票番号の開始値。前回出力分の続きから採番したい場合に指定
 */
function buildPcaCsv(entries, mode, voucherStart) {
  return new PcaInvoiceWindow().buildCsv(entries, mode, voucherStart);
}

/** ⚙ マスタ一式の取得（画面初期化用） */
function getMasters() {
  return MasterStore.loadAll();
}

/** ⚙ 社内管理科目（PL（詳細）D〜E列）の貼り付け取り込み */
function importInternalAccounts(text) {
  return MasterStore.importInternalAccounts(text);
}

/** ⚙ PCA勘定科目マスタの貼り付け取り込み */
function importPcaAccounts(text) {
  return MasterStore.importPcaAccounts(text);
}

/** ⚙ 部門マスタの貼り付け取り込み */
function importDepartments(text) {
  return MasterStore.importDepartments(text);
}

/** ⚙ 指定マスタを初期値（AccountMaster.gs のシード）に戻す */
function resetMaster(name) {
  MasterStore.reset(name);
  return MasterStore.loadAll();
}

/** ⚙ 画面の表UIで編集したマスタ一覧を保存する（行の追加・修正・削除） */
function saveMasterList(name, list) {
  return MasterStore.saveMasterList(name, list);
}

/** ⚙ 仕分けルールを1件追加（窓5の未確定リストからその場で登録） */
function addKeywordRule(rule) {
  return MasterStore.addKeywordRule(rule);
}

/** ⚙ 社内管理科目を1件追加・更新（未確定リストから新しい科目を作る） */
function addInternalAccount(item) {
  return MasterStore.addInternalAccount(item);
}

/**
 * 🧾 窓5：マスタを更新したあと、解析済みの明細をもう一度判定し直す。
 * AIへの再問い合わせは行わないため、追加のAPI料金も待ち時間も発生しません。
 */
function reclassifyEntries(entries) {
  return new PcaInvoiceWindow().reclassify(entries);
}

/** 📗 Googleスプレッドシートとして出力し、URLを返す */
function exportToSpreadsheet(payload) {
  return new SheetExporter().toSpreadsheet(payload);
}

/** 📘 Excel(.xlsx) として出力し、Base64を返す（ブラウザ側でダウンロード） */
function exportToExcel(payload) {
  return new SheetExporter().toExcelBase64(payload);
}

/**
 * 🌐 画面ログの翻訳・要約サポート（エラーログ表示専用のユーティリティ）
 * ここで例外が発生しても null を返すだけで、アプリ本体の動作・
 * パフォーマンスには一切影響しません。
 */
function translateAndSummarizeLog(originalMessage) {
  return new LogTranslator().translate(originalMessage);
}


// ====================================================================
// CsvUtil：CSVの安全な組み立て（文字化け・列ズレ対策の中核）
// ------------------------------------------------------------------
// V7.00 では row.join(',') をそのまま使っていたため、摘要に「,」や
// 改行が含まれると列が丸ごとズレて、画面上は文字化けのように見えて
// いました。ここで RFC4180 準拠のクォート処理を行います。
// ====================================================================
class CsvUtil {

  /**
   * セル値のクリーニング。
   * ・Excel由来の "_x000D_"、制御文字、置換文字(U+FFFD) を除去
   * ・改行は半角スペースに畳む（1明細＝1行を保証するため）
   * ・全角スペースは半角スペースへ
   */
  static sanitize(value) {
    if (value === null || value === undefined) return '';
    let s = String(value);
    s = s.replace(/_x000D_/g, '');
    s = s.replace(/\uFFFD/g, '');            // 置換文字（文字化けの痕跡）を除去
    s = s.replace(/[\r\n]+/g, ' ');          // 改行は半角スペースへ畳む
    s = s.replace(/[\u0000-\u001F\u007F]/g, ''); // 制御文字を除去
    s = s.replace(/\u3000/g, ' ');           // 全角スペース→半角スペース
    s = s.replace(/ {2,}/g, ' ');
    return s.trim();
  }

  /** 1セルをCSVフィールドへ。「,」「"」を含む場合のみクォートします。 */
  static field(value) {
    const s = CsvUtil.sanitize(value);
    if (s.indexOf(',') !== -1 || s.indexOf('"') !== -1) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }

  /** 1行分の配列をCSV行文字列へ。 */
  static row(values) {
    return (values || []).map(CsvUtil.field).join(',');
  }

  /** ヘッダー＋明細からCSV全文を組み立てます。 */
  static build(headers, rows) {
    const lines = [];
    if (headers && headers.length) lines.push(CsvUtil.row(headers));
    (rows || []).forEach(function (r) { lines.push(CsvUtil.row(r)); });
    return lines.join('\n');
  }
}


// ====================================================================
// ModelRegistry：利用可能なGeminiモデルのホワイトリスト管理
// （不正な文字列がそのままURLに使われないようにするための安全弁）
// ====================================================================
class ModelRegistry {
  static get DEFAULT() { return 'gemini-3.5-flash-lite'; }
  static get ALLOWED() { return ['gemini-3.5-flash-lite', 'gemini-3.6-flash']; }

  static resolve(modelName) {
    if (modelName && ModelRegistry.ALLOWED.indexOf(modelName) !== -1) {
      return modelName;
    }
    // 未指定・不正値の場合は、旧バージョンと同じデフォルトモデルにフォールバック
    return ModelRegistry.DEFAULT;
  }
}


// ====================================================================
// GeminiApiClient：Geminiとの通信のみを担当する共通クラス
// 各窓のプロンプト内容・出力フォーマットには一切関与しません。
// ====================================================================
class GeminiApiClient {
  constructor(modelName) {
    this.apiKey = PropertiesService.getScriptProperties().getProperty('USER_GEMINI_API_KEY');
    this.model = ModelRegistry.resolve(modelName);
  }

  hasApiKey() {
    return !!this.apiKey;
  }

  _endpoint() {
    return 'https://generativelanguage.googleapis.com/v1beta/models/' +
      encodeURIComponent(this.model) + ':generateContent?key=' + encodeURIComponent(this.apiKey);
  }

  /**
   * 一時的なエラー（429/500/503）は指数バックオフで最大3回まで再試行します。
   * 恒久的なエラーはその場で例外を投げます。
   */
  _fetchWithRetry(options) {
    const maxAttempts = 3;
    let lastText = '';
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      const response = UrlFetchApp.fetch(this._endpoint(), options);
      const code = response.getResponseCode();
      lastText = response.getContentText();
      if (code === 200) return lastText;
      if (code === 429 || code === 500 || code === 503) {
        if (attempt < maxAttempts) {
          Utilities.sleep(1000 * Math.pow(2, attempt - 1));
          continue;
        }
      }
      let message = 'Gemini APIエラー (HTTP ' + code + ')';
      try {
        const errJson = JSON.parse(lastText);
        if (errJson && errJson.error && errJson.error.message) message = errJson.error.message;
      } catch (parseErr) {
        message += ': ' + lastText.substring(0, 300);
      }
      throw new Error(message);
    }
    throw new Error('Gemini APIが混雑しています。しばらく待って再実行してください。');
  }

  /**
   * レスポンスから本文テキストを取り出す。
   * ⚠ V7.00 では parts[0] しか読んでいなかったため、応答が複数パートに
   *   分割された場合にJSONが途中で切れ、パース失敗＝文字化けの原因に
   *   なっていました。ここで全パートを連結します。
   */
  _extractText(resText) {
    const jsonRes = JSON.parse(resText);

    if (jsonRes.promptFeedback && jsonRes.promptFeedback.blockReason) {
      throw new Error('AIが安全フィルタによって応答を拒否しました（理由: ' +
        jsonRes.promptFeedback.blockReason + '）。ファイルを変えて再実行してください。');
    }

    const candidate = jsonRes.candidates && jsonRes.candidates[0];
    if (!candidate) throw new Error('AIから有効な応答が返りませんでした。');

    if (candidate.finishReason === 'MAX_TOKENS') {
      throw new Error('データ量が多すぎて応答が途中で打ち切られました。ファイルを分割して再実行してください。');
    }

    const parts = (candidate.content && candidate.content.parts) || [];
    const text = parts.map(function (p) { return p.text || ''; }).join('');
    if (!text) throw new Error('AIの応答が空でした。');
    return text;
  }

  /**
   * ファイル(画像/PDF) ＋ プロンプトを送信し、レスポンス内のJSON部分を
   * 抽出して返す。窓1・窓2・窓3・窓5が共通で利用する。
   */
  generateJson(base64Data, mimeType, prompt) {
    const payload = {
      'contents': [
        {
          'parts': [
            { 'text': prompt },
            {
              'inline_data': {
                'mime_type': mimeType,
                'data': base64Data
              }
            }
          ]
        }
      ],
      // 応答をJSONに固定。Markdownの ```json 枠が混入しなくなります。
      'generationConfig': {
        'responseMimeType': 'application/json',
        'temperature': 0
      }
    };

    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    const resultText = this._extractText(this._fetchWithRetry(options));

    try {
      return JSON.parse(resultText);
    } catch (e) {
      // responseMimeType が効かない環境向けのフォールバック（V7.00 と同じ挙動）
      const jsonMatch = resultText.match(/\[[\s\S]*\]|\{[\s\S]*\}/);
      if (!jsonMatch) {
        throw new Error('データが見つかりませんでした。AIの返答: ' + resultText.substring(0, 300));
      }
      return JSON.parse(jsonMatch[0]);
    }
  }

  /**
   * テキストのみのシンプルな生成（ログ翻訳・要約用の軽量ユーティリティ）。
   * ファイル解析（窓1〜3・窓5）とは完全に独立しており、失敗時は null を返す。
   */
  generateText(prompt) {
    const payload = { 'contents': [{ 'parts': [{ 'text': prompt }] }] };
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(this._endpoint(), options);
    if (response.getResponseCode() !== 200) return null;

    try {
      return this._extractText(response.getContentText());
    } catch (e) {
      return null;
    }
  }
}


// ====================================================================
// SharedDateRules：窓1・窓2「のみ」が共有する和暦・西暦変換ルール。
// 窓3・窓5は独自の日付ルールを持つため、あえてここには含めていません。
// ====================================================================
class SharedDateRules {
  static get PROMPT() {
    return `
＜日付の和暦・西暦変換の標準ルール＞
本日は西暦【2026年（令和8年）】です。
通常、ビジネスデータは現行（2025年〜2026年等）として処理されますが、後述の「追加指示」で特定の年号・西暦指定がある場合は本ルールを無視し、指示を絶対最優先してください。

通帳の左端に印字されている「08-02-05」「08-03-10」などの文字列は、必ず【YY-MM-DD】の並び順（年-月-日）です。
AIの勝手な思い込みで、先頭の「08」を「8月」や「2024年の8」などと位置をズラして解釈することは厳禁とします。
以下の基本構造をあてはめて4桁の西暦（YYYY/MM/DD）を計算してください。

【構造公式】
左端の文字列が「AA-BB-CC」または「AA.BB.CC」の場合：
・「AA」＝【年（和暦）】
・「BB」＝【月】
・「CC」＝【日】

【和暦「AA」の変換確定ルール（指示がない場合のデフォルト）】
1. AA が「08」の場合：
   ・これは「令和8年」です。
   ・デフォルトは【2026年】とします（計算式：2019 + 8 = 2026）。
   ・「2024/8/2」のように「8」を月にスライドさせたりすることは致命的なバグです。必ず【2026/BB/CC】（例：2026/02/05）と出力してください。

2. AA が「07」の場合：
   ・これは「令和7年」です。デフォルト【2025年】とします（計算式：2019 + 7 = 2025）。

3. AA が「26 〜 31」の場合：
   ・これは「平成」です。必ず【1988 + AA】で計算してください。

【行間の整合性チェック】
・同一ファイル内の日付は、すべて同じ「年（AA）」を推移します。数行だけ別の年にジャンプしたり、月日がねじれることはありません。必ず前後の日付と整合させてください。
`;
  }

  /** ユーザー指示を最優先にするためのプロンプト注入ラッパー（窓1・窓2共通） */
  static buildFinalPrompt(basePrompt, customPrompt) {
    if (!customPrompt) {
      return basePrompt;
    }
    return `【🚨最優先・最上位絶対命令🚨】
現在、システムに以下の追加指示（ユーザーの直接命令）が与えられています。
『${customPrompt}』
システム側のいかなるデフォルトルールや日付計算ロジック、和暦推測よりも、この『${customPrompt}』の指示内容を100%最優先してください。もし競合や矛盾がある場合は、デフォルトルールを完全に無視（オーバーライド）して指示内容に従ったCSVデータを構築してください。
------------------
${basePrompt}
------------------
【🚨もう一度確認してください：最重要命令🚨】
『${customPrompt}』
上記の指示内容が完璧に反映されたJSONを出力してください。`;
  }
}


// ====================================================================
// 🏦 窓1：通帳・現金 専用クラス
// このクラス内の編集は窓2・窓3・窓4・窓5に一切影響しません。
// ====================================================================
class CashWindow {
  buildPrompt(customPrompt) {
    const basePrompt = `【最重要：通帳・現金データ抽出 ＆ 残高によるセルフ計算整合性チェック】
提出されたファイルから取引明細を正確に抽出し、指定のJSONオブジェクトのみを返してください。

${SharedDateRules.PROMPT}

＜最重要ミッション：残高データを使用した金額ダブルチェック＞
1. 【残高ステータスの有無判定】:
   アップロードされた資料内に「残高」「差引残高」などの列や記載が存在するかどうかを判定してください。存在する場合は has_balance を true、ない場合は false とします。

2. 【残高整合性チェックロジック（has_balanceがtrueの場合にのみ適用）】:
   ・読み取った各行について、次の数式が必ず成立しているかを「AI自身で検証（セルフチェック）」してください：
     「（前行の）残高」 + 「（今行の）入金」 - 「（今行の）出金」 ＝ 「（今行の）残高」
   ・もしこの計算式の整合性が合わないと判定した場合、周囲の数字パターンや印字状態から合理的・論理的に判断し、「もっとも正しいと思われる入金・出金・残高の数値」へと自動的に補正してください。

・値がない項目（入出金がない空欄）は必ず null にすること。「0」は入れないでください。
・読み取れない文字を憶測で埋めたり、意味不明な記号列で埋めることは禁止です。読めない場合は null としてください。
・余計な挨拶や解説、Markdownの枠（\`\`\`jsonなど）は一切出力せず、純粋なJSONオブジェクトのみを返してください。

＜出力フォーマット＞
{
  "has_balance": true,
  "records": [
    {
      "date": "YYYY/MM/DD",
      "in": 10000,
      "out": null,
      "balance": 250000,
      "description": "お振込 ヤマダ"
    }
  ]
}`;

    return SharedDateRules.buildFinalPrompt(basePrompt, customPrompt);
  }

  process(base64Data, mimeType, customPrompt, modelName) {
    const client = new GeminiApiClient(modelName);
    if (!client.hasApiKey()) {
      throw new Error('APIキーが設定されていません。画面右上の『Gemini APIキー設定』から登録してください。');
    }

    try {
      const prompt = this.buildPrompt(customPrompt);
      const data = client.generateJson(base64Data, mimeType, prompt);
      return this.formatOutput(data);
    } catch (e) {
      throw new Error('解析失敗: ' + e.message);
    }
  }

  /** 窓1専用の出力整形。窓2・窓3・窓5のフォーマット処理とは完全に独立しています。 */
  formatOutput(data) {
    const hasBalance = data.has_balance === true;
    const records = data.records || [];
    const headers = hasBalance
      ? ['西暦取引年月日', '入金', '出金', '残高', '摘要']
      : ['西暦取引年月日', '入金', '出金', '摘要'];
    const rows = [];

    records.forEach(function (row) {
      const inVal = (row.in && row.in !== 0) ? row.in : '';
      const outVal = (row.out && row.out !== 0) ? row.out : '';
      const balanceVal = (row.balance !== null && row.balance !== undefined) ? row.balance : '';
      const desc = CsvUtil.sanitize(row.description);

      rows.push(hasBalance
        ? [row.date, inVal, outVal, balanceVal, desc]
        : [row.date, inVal, outVal, desc]);
    });

    return { csv: CsvUtil.build(headers, rows), headers: headers, rows: rows };
  }
}


// ====================================================================
// 💳 窓2：クレジット（スマート取引取込用） 専用クラス
// このクラス内の編集は窓1・窓3・窓4・窓5に一切影響しません。
// ====================================================================
class SmartCreditWindow {
  buildPrompt(customPrompt) {
    const basePrompt = `【最重要：スマート取引取込用 クレジットカード明細データ抽出】
必ず、日付、金額、店舗名（摘要）、推定使用目的の4つから構成される以下のJSONオブジェクトで出力してください。

${SharedDateRules.PROMPT}

＜抽出時の注意＞
・日付は必ず「YYYY/MM/DD」形式の西暦にすること。
・金額は純粋な数値（カンマなし）のみを抽出すること。
・店舗名（description）には、日付（MM/DDなど）の文字情報を「絶対に」含めないでください。純粋な利用店舗名や内容のみにすること。
・目的欄（purpose）には、抽出した店舗名や利用内容から連想される「一般的なビジネス上の使用目的や勘定科目の目安（例：消耗品費、旅費交通費、飲食費、通信費、広告宣伝費など）」をAIの知識から合理的に推測して記載してください。判断できない場合は空文字 "" としてください（意味のない記号列で埋めないこと）。
・余計な挨拶や解説、Markdownの枠（\`\`\`jsonなど）は一切出力せず、純粋なJSONオブジェクトのみを返してください。

＜出力フォーマット＞
{
  "records": [
    {
      "date": "YYYY/MM/DD",
      "amount": 1280,
      "description": "店舗名や利用内容",
      "purpose": "消耗品費や飲食費"
    }
  ]
}`;

    return SharedDateRules.buildFinalPrompt(basePrompt, customPrompt);
  }

  process(base64Data, mimeType, customPrompt, modelName) {
    const client = new GeminiApiClient(modelName);
    if (!client.hasApiKey()) {
      throw new Error('APIキーが設定されていません。画面右上の『Gemini APIキー設定』から登録してください。');
    }

    try {
      const prompt = this.buildPrompt(customPrompt);
      const data = client.generateJson(base64Data, mimeType, prompt);
      return this.formatOutput(data);
    } catch (e) {
      throw new Error('解析失敗: ' + e.message);
    }
  }

  /** 窓2専用の出力整形。窓1・窓3・窓5のフォーマット処理とは完全に独立しています。 */
  formatOutput(data) {
    const records = data.records || [];
    const headers = ['西暦年月日', '金額', '摘要', '推定使用目的'];
    const rows = [];

    records.forEach(function (row) {
      const purposeVal = row.purpose ? CsvUtil.sanitize(row.purpose) : '';
      rows.push([row.date, row.amount, CsvUtil.sanitize(row.description), purposeVal]);
    });

    return { csv: CsvUtil.build(headers, rows), headers: headers, rows: rows };
  }
}


// ====================================================================
// 💳 窓3：クレジット（PC版弥生インポート用 中間ファイル） 専用クラス
// ver1（V5.00）仕様を維持。窓1・窓2とは日付ルール・出力形式が異なるため
// 独自実装とし、SharedDateRules には依存させていません。
// このクラス内の編集は窓1・窓2・窓4・窓5に一切影響しません。
// ====================================================================
class CreditWindow {
  buildPrompt(customPrompt) {
    let prompt = `【最重要：PC版弥生仕訳インポート用 クレジットカード明細データ抽出】
必ず、日付、金額、店舗名（適用）の3つから構成される以下の単一金額（amount）形式のJSON配列で出力してください。
・「2026」や西暦が離れた場所に記載されていても、結合して必ず「YYYY/MM/DD」形式の西暦にすること。
・金額に含まれる支払回数（1回など）の文字は除外し、純粋な数値のみを抽出すること。
・読み取れない文字を憶測や記号列で埋めることは禁止です。読めない場合は空文字 "" としてください。
・余計な挨拶や解説、Markdownの枠（\`\`\`jsonなど）は一切出力せず、純粋なJSON配列のみを返してください。

フォーマット：
[{"date": "YYYY/MM/DD", "amount": 1280, "description": "店舗名や利用内容"}]`;

    if (customPrompt) {
      prompt += `\n【ユーザーからの追加指示（最優先）】:${customPrompt}`;
    }

    return prompt;
  }

  process(base64Data, mimeType, customPrompt, modelName) {
    const client = new GeminiApiClient(modelName);
    if (!client.hasApiKey()) {
      throw new Error('APIキーが設定されていません。画面右上の『Gemini APIキー設定』から登録してください。');
    }

    try {
      const prompt = this.buildPrompt(customPrompt);
      const data = client.generateJson(base64Data, mimeType, prompt);
      return this.formatOutput(data);
    } catch (e) {
      throw new Error('解析失敗: ' + e.message);
    }
  }

  /**
   * 窓3専用の出力整形（純粋なCSVテキスト文字列を直接返却）。
   * 列構成はV7.00と同一。摘要のクォート処理のみ追加しています。
   */
  formatOutput(data) {
    const records = Array.isArray(data) ? data : (data.records || []);
    const headers = ['西暦年月日', '金額', '摘要（日付入り）'];
    const rows = [];

    records.forEach(function (row) {
      let mmdd = '';
      try {
        const parts = String(row.date).split('/');
        mmdd = parts.length >= 3 ? parseInt(parts[1], 10) + '/' + parseInt(parts[2], 10) : '??/??';
      } catch (err) {
        mmdd = '??/??';
      }
      rows.push([row.date, row.amount, mmdd + ' ' + CsvUtil.sanitize(row.description)]);
    });

    return CsvUtil.build(headers, rows) + '\n';
  }
}


// ====================================================================
// 🧾 窓5：請求書・領収書 → PCA会計 仕訳変換 専用クラス
// ------------------------------------------------------------------
// ・窓1〜窓4のクラスには一切依存せず、影響も与えません。
// ・複数ファイルを解析して蓄積し、最後に1つのCSVへまとめます。
// ・科目が不明な場合は【必ず空欄】。推測での当てはめはしません。
// ====================================================================
class PcaInvoiceWindow {

  /** 摘要（社内向け）のヘッダー。実データ列は buildInternalRows と対応。 */
  static get INTERNAL_HEADERS() {
    return [
      '伝票日付', '伝票番号',
      '借方科目コード', '借方科目名', '借方補助コード', '借方補助名',
      '借方部門コード', '借方部門名', '借方税区分名', '借方金額',
      '摘要文', '大分類', '小分類', '取引先', '元ファイル名', '要確認メモ'
    ];
  }

  /** 会計事務所提出用（PCA仕訳インポート）のヘッダー。 */
  static get OFFICE_HEADERS() {
    return [
      '伝票日付', '伝票番号',
      '借方科目コード', '借方科目名', '借方補助コード', '借方補助名',
      '借方部門コード', '借方部門名', '借方税区分名', '借方金額',
      '貸方科目コード', '貸方科目名', '貸方補助コード', '貸方補助名',
      '貸方部門コード', '貸方部門名', '貸方税区分名', '貸方金額',
      '摘要文'
    ];
  }

  /** 自社名（請求先）。ここに挙げた名称は「取引先」として採用しません。 */
  static get OWN_COMPANY_NAMES() {
    return ['株式会社エルティーアール', 'エルティーアール', 'LTR'];
  }

  buildPrompt(customPrompt, masters) {
    const internal = (masters.internalAccounts || []);

    // 大分類／小分類のペア一覧をプロンプトへ渡し、マスタ外の科目名を作らせない。
    // 「版権元RY」のように同名の小分類が複数の大分類に存在するため、
    // 必ずペアで選ばせます。
    const pairList = internal.map(function (a) {
      return '・大分類「' + (a.major || '') + '」／小分類「' + (a.minor || '') + '」'
        + (a.section ? '（' + a.section + '）' : '');
    }).join('\n');

    const deptList = (masters.departments || []).map(function (d) {
      return '・' + d.name;
    }).join('\n');

    let prompt = `【最重要：請求書・領収書からの 会計仕訳データ抽出（PCA会計向け）】
アップロードされた請求書・領収書を読み取り、下記JSONオブジェクトのみを返してください。
印刷された請求書だけでなく、手書きの請求書やスキャン画像も対象です。

＜絶対に守る出力ルール＞
1. 【判断できないものは必ず null】
   ・勘定科目が判断できない、請求書に手がかりが無い場合は、必ず null を返してください。
   ・「不明」「？」「N/A」などの文字列や、読めない文字を記号で埋めた文字列を返すことは【厳禁】です。
   ・後から人間が手入力して埋めるため、空（null）であることが正解です。推測で埋めないでください。
2. 【科目は下記リストの「大分類＋小分類のペア」からのみ選ぶ】
   ・major_category と minor_category は、必ず同じ1行のペアをそのまま転記してください。
   ・リストに無い科目名を新しく作ってはいけません。当てはまるものが無ければ両方とも null。
   ・「版権元RY」のように同じ小分類名が複数の大分類にあります。どの大分類か特定できない場合は
     両方とも null にしてください（片方を勝手に選ばないこと）。
3. 【数字は絶対に連結しない】
   ・金額はカンマ・通貨記号・「円」を除いた純粋な数値のみ。
   ・請求書番号・電話番号・登録番号・口座番号・郵便番号を金額として拾わないでください。
   ・別々のセルに書かれた数字を1つの数値としてつなげることは【絶対禁止】です
     （例：9,750,000 と 975,000 を 9750000975000 のようにつなげてはいけません）。
   ・1件の金額が 100億円を超えることは通常ありません。そうなった場合は読み取りを間違えています。
4. Markdownの枠（\`\`\`jsonなど）や挨拶・解説は一切出力せず、純粋なJSONオブジェクトのみを返してください。

＜取引先（vendor）の判定＞
・vendor には【請求元＝お金を受け取る側】の会社名を入れてください。
・「${PcaInvoiceWindow.OWN_COMPANY_NAMES.join('」「')}」は当社（＝請求先・支払う側）です。
  「御中」「様」が付いている宛先は請求先なので、絶対に vendor に入れないでください。
・請求元は多くの場合、右上・右下の社判（角印）や振込先口座の名義の近くに記載されています。

＜日付（date）の判定＞
・必ず「YYYY/MM/DD」形式の西暦で返してください。
・和暦の場合は西暦へ変換してください。「令和」の表記が無く数字だけの場合も和暦とみなします。
    「8年6月30日」「R8.6.30」「令和8年6月30日」 → いずれも 2026/06/30
    （計算式：令和 = 2018 + 和暦年。令和8年 = 2026年）
・「請求日」「発行日」を優先します。無ければ「締切日」「締日」「売上日付」を使ってください。
  「支払期限」「お支払期日」は使わないでください。
・年がどこにも書かれていない場合のみ null にしてください。

＜金額（net_amount / tax_amount / total_amount）の判定＞
・net_amount   ＝ 税抜金額（「税抜額」「小計」「課税対象額」「税率別内訳の税抜金額」など）
・tax_amount   ＝ 消費税額（「消費税」「消費税等」「税」など）
・total_amount ＝ 税込合計（「合計（税込）」「今回請求額」「御請求額」「税込額」など）
・【最重要】「今回請求額」「税込」と書かれた金額は税込合計です。税抜金額として扱わないでください。
・「税率別内訳」「税率別税抜合計」の表がある場合は、そこの税抜金額と消費税額を最優先で使ってください。
・net_amount + tax_amount = total_amount が必ず成り立つはずです。合わない場合は読み直してください。
・読み取れなかった項目は null にしてください（他の値から計算して埋めないこと。計算はシステム側で行います）。

＜明細の分け方（records の作り方）＞
・請求書の品名（明細行）ごとに1レコードを作ってください。
・複数ページある場合、表紙の合計だけでなく明細ページの各行を読み取ってください。
  明細ページがある場合は、表紙の合計行をレコードにしないでください（二重計上になります）。
・明細行ごとの消費税額が書かれていない場合は、その行の tax_amount は null にしてください。
・「小計」「合計」「消費税」「繰越額」だけの行はレコードにしないでください。

＜選択可能な 大分類 / 小分類 の一覧（この中からペアで選ぶ）＞
${pairList || '（マスタ未設定。major_category と minor_category はすべて null で返してください）'}

＜選択可能な 部門名の一覧（該当が無ければ null）＞
${deptList || '（部門マスタ未設定。null で返してください）'}

＜税率＞
・tax_rate は 10 または 8 のみ。軽減税率（「※」印や「軽」の表記）の場合は is_reduced を true。
・税率の記載が読み取れない場合は null にしてください。

＜出力フォーマット＞
{
  "records": [
    {
      "date": "YYYY/MM/DD",
      "vendor": "請求元の会社名",
      "invoice_no": "請求書番号（無ければ null）",
      "net_amount": 231000,
      "tax_amount": 23100,
      "total_amount": 254100,
      "tax_rate": 10,
      "is_reduced": false,
      "major_category": "共通",
      "minor_category": "デザイン費",
      "department": "共通部門",
      "description": "デフォルメイラスト制作費"
    }
  ]
}`;

    if (customPrompt) {
      prompt = `【🚨最優先・最上位絶対命令🚨】
以下の追加指示（ユーザーの直接命令）を、システム側のいかなるデフォルトルールよりも100%最優先してください。
『${customPrompt}』
ただし「判断できないものは null で返す」という原則だけは常に維持してください。
------------------
${prompt}
------------------
【🚨もう一度確認：最重要命令🚨】
『${customPrompt}』`;
    }

    return prompt;
  }

  /**
   * 1ファイル分を解析し、マスタで正規化した「明細オブジェクトの配列」を返します。
   * ここではCSVにはせず、画面側で複数ファイル分を蓄積してから
   * buildCsv() でまとめて1つのCSVにします。
   */
  process(base64Data, mimeType, customPrompt, modelName, sourceFileName) {
    const client = new GeminiApiClient(modelName);
    if (!client.hasApiKey()) {
      throw new Error('APIキーが設定されていません。画面右上の『Gemini APIキー設定』から登録してください。');
    }

    const masters = MasterStore.loadAll();
    const resolver = new AccountResolver(masters);

    let data;
    try {
      const prompt = this.buildPrompt(customPrompt, masters);
      data = client.generateJson(base64Data, mimeType, prompt);
    } catch (e) {
      throw new Error('解析失敗: ' + e.message);
    }

    const records = (data && data.records) ? data.records : (Array.isArray(data) ? data : []);
    const fileName = CsvUtil.sanitize(sourceFileName);
    const self = this;

    return records.map(function (rec) {
      return self._normalize(rec, resolver, fileName);
    });
  }

  /**
   * マスタを更新したあとに、解析済みの明細をもう一度判定し直します。
   * 各明細にはAIの生出力（raw）を保持してあるので、Geminiへの再問い合わせは
   * 不要です（追加料金も待ち時間も発生しません）。
   *
   * raw が無い明細（画面で手入力した行など）は、そのまま返して壊しません。
   */
  reclassify(entries) {
    const resolver = new AccountResolver(MasterStore.loadAll());
    const self = this;
    return (entries || []).map(function (entry) {
      if (!entry || !entry.raw) return entry;
      return self._normalize(entry.raw, resolver, entry.sourceFile);
    });
  }

  /**
   * AIの生出力を、マスタに突き合わせた確定値へ変換します。
   * マスタで解決できなかった項目は必ず '' （空欄）になります。
   */
  _normalize(rec, resolver, fileName) {
    const warnings = [];
    let appliedRule = '';

    // --- 科目の決定：社内ルール > AI判断。どちらも駄目なら空欄。 ---
    // 科目は【大分類＋小分類のペア】で確定させます。
    // 取引先と摘要を分けて渡すことで、「この取引先はこの科目」という
    // 摘要に情報が無い請求書にも対応できます。
    const keywordHit = resolver.applyKeywordRule(rec.description, rec.vendor);
    const minorName = keywordHit ? keywordHit.minor : CsvUtil.sanitize(rec.minor_category);
    const majorName = keywordHit ? keywordHit.major : CsvUtil.sanitize(rec.major_category);
    const internal = resolver.resolveInternal(minorName, majorName);

    if (keywordHit && internal) {
      const hitRule = keywordHit.rule || {};
      const label = [hitRule.vendorKeyword, hitRule.keyword].filter(Boolean).join(' + ');
      if (label) appliedRule = '社内ルール「' + label + '」を適用';
    }

    if (!internal && minorName && resolver.isAmbiguousMinor(minorName)) {
      // 例：「版権元RY」は飲食・物販・予約の3つに存在。大分類が分からないと確定できない。
      warnings.push('小分類「' + minorName + '」は複数の大分類に存在するため確定できません');
    }

    // --- PCA勘定科目の決定：社内科目からの紐付けのみ。無ければ空欄。 ---
    let pcaAccount = null;
    if (internal && internal.pcaCode) {
      pcaAccount = resolver.resolvePcaAccount(internal.pcaCode);
    }
    if (!pcaAccount && internal) {
      warnings.push('社内科目「' + internal.minor + '」にPCA勘定科目が紐付いていません');
    }

    const major = internal ? internal.major : '';
    const minor = internal ? internal.minor : '';

    // --- 金額：読めた値のみを採用し、足りない分だけ算術で補完する ---
    let net = PcaInvoiceWindow._num(rec.net_amount);
    let tax = PcaInvoiceWindow._num(rec.tax_amount);
    let total = PcaInvoiceWindow._num(rec.total_amount);

    if (net === null && total !== null && tax !== null) net = total - tax;
    if (tax === null && total !== null && net !== null) tax = total - net;
    if (total === null && net !== null && tax !== null) total = net + tax;

    // 税抜＋税額＝税込 が成立しない場合、税抜と税額（内訳側）を正として税込を再計算。
    // AIが「今回請求額（税込）」を税抜として拾ってしまう事故をここで吸収します。
    if (net !== null && tax !== null && total !== null && net + tax !== total) {
      warnings.push('税抜+消費税(' + (net + tax) + ')と税込(' + total + ')が不一致のため税込を再計算しました');
      total = net + tax;
    }

    // --- 税区分：税率と軽減税率フラグからマスタ名を組み立てる ---
    const rate = PcaInvoiceWindow._num(rec.tax_rate);
    let taxLabel = '';
    if (rate === 10) taxLabel = '仕入(課)10%';
    else if (rate === 8) taxLabel = rec.is_reduced ? '仕入(課)8%軽' : '仕入(課)8%';
    const taxClass = resolver.resolveTax(taxLabel) || '';

    // --- 部門：マスタに無ければ空欄 ---
    const dept = resolver.resolveDepartment(rec.department);

    // --- 補助科目：取引先名がマスタにあれば採用、無ければ空欄 ---
    const sub = pcaAccount ? resolver.resolveSubAccount(pcaAccount.code, rec.vendor) : null;

    // --- 取引先：自社名（＝請求先）を拾ってしまっていたら空欄にする ---
    let vendor = CsvUtil.sanitize(rec.vendor);
    if (vendor && PcaInvoiceWindow._isOwnCompany(vendor)) {
      warnings.push('取引先として自社名が読み取られたため空欄にしました');
      vendor = '';
    }

    const date = PcaInvoiceWindow._normalizeDate(rec.date);
    if (!date && rec.date) warnings.push('日付を西暦に確定できませんでした');

    return {
      date:        date,
      vendor:      vendor,
      invoiceNo:   CsvUtil.sanitize(rec.invoice_no),
      description: CsvUtil.sanitize(rec.description),
      major:       major,
      minor:       minor,
      accountCode: pcaAccount ? String(pcaAccount.code) : '',
      accountName: pcaAccount ? pcaAccount.name : '',
      subCode:     sub ? String(sub.code) : '',
      subName:     sub ? sub.name : '',
      deptCode:    dept ? String(dept.code) : '',
      deptName:    dept ? dept.name : '',
      taxClass:    taxClass,
      netAmount:   net === null ? '' : net,
      taxAmount:   tax === null ? '' : tax,
      totalAmount: total === null ? '' : total,
      sourceFile:  fileName,
      appliedRule: appliedRule,
      warnings:    warnings.join(' / '),
      // マスタを更新したあとに再判定できるよう、AIの生出力を保持しておきます。
      // これにより「あとで科目を設定 → 再判定」がAI再解析なしで完了します。
      raw:         rec
    };
  }

  /**
   * 科目が未確定（空欄）の明細を、後から設定するための一覧として抽出します。
   * 同じ「取引先＋摘要」はまとめ、件数の多い順に並べます
   * （よく出るものから設定していけるように）。
   */
  static collectUnresolved(entries) {
    const groups = {};
    (entries || []).forEach(function (e) {
      if (!e || e.accountCode) return;   // 科目が確定済みの明細は対象外
      const vendor = e.vendor || '';
      const desc = e.description || '';
      const key = vendor + '||' + desc;
      if (!groups[key]) {
        groups[key] = {
          vendor: vendor,
          description: desc,
          // ルール化するときの初期キーワード（摘要を優先、無ければ取引先）
          suggestedKeyword: desc || vendor,
          count: 0,
          amount: 0,
          files: [],
          reason: e.warnings || (e.minor ? 'PCA勘定科目が未紐付け' : '勘定科目が判定できませんでした'),
          currentMajor: e.major || '',
          currentMinor: e.minor || ''
        };
      }
      const g = groups[key];
      g.count++;
      g.amount += (typeof e.totalAmount === 'number') ? e.totalAmount : 0;
      if (g.files.indexOf(e.sourceFile) === -1 && e.sourceFile) g.files.push(e.sourceFile);
    });

    return Object.keys(groups).map(function (k) { return groups[k]; })
      .sort(function (a, b) { return b.count - a.count || b.amount - a.amount; });
  }

  /** 自社名（請求先）かどうか。取引先として採用しないための判定。 */
  static _isOwnCompany(name) {
    const key = AccountResolver.normalize(name);
    return PcaInvoiceWindow.OWN_COMPANY_NAMES.some(function (own) {
      const ownKey = AccountResolver.normalize(own);
      return ownKey && key.indexOf(ownKey) !== -1;
    });
  }

  /** 1件の金額として現実的な上限。これを超えたら読み取り失敗とみなします。 */
  static get MAX_AMOUNT() { return 1e11; }

  /**
   * 金額の数値化。
   * ⚠ 桁の連結事故（別セルの数字がつながって 833,111,791,313,101 のような
   *   巨大な値になる）を検出したら、推測で直さず null（＝空欄）にします。
   */
  static _num(value) {
    if (value === null || value === undefined || value === '') return null;
    const cleaned = String(value).replace(/[,，\s　円¥￥]/g, '');
    const n = Number(cleaned.replace(/[^0-9.\-]/g, ''));
    if (!isFinite(n)) return null;
    if (Math.abs(n) > PcaInvoiceWindow.MAX_AMOUNT) return null;
    return n;
  }

  /** 元号ごとの「西暦 = 元号元年の前年」。和暦N年 → BASE + N。 */
  static get ERA_BASE() {
    return { '令和': 2018, 'R': 2018, '平成': 1988, 'H': 1988, '昭和': 1925, 'S': 1925 };
  }

  /**
   * "YYYY/MM/DD" へ寄せる。判定できない場合は空欄（推測しない）。
   *
   * 手書き請求書では「8年6月30日」のように元号を省いた和暦がよく使われるため、
   * 1〜2桁の年は和暦（令和）として西暦に変換します。
   *   例：8年6月30日 → 2026/06/30（2018 + 8）
   * 4桁の年はそのまま西暦として扱います。
   */
  static _normalizeDate(value) {
    if (!value) return '';
    const raw = CsvUtil.sanitize(value);
    const pad = function (n) { return ('0' + n).slice(-2); };

    // 元号表記あり（令和8年6月30日 / R8.6.30 / 令和08-06-30）
    const eraMatch = raw.match(/(令和|平成|昭和|[RHS])\s*(\d{1,2})\s*[年.\-\/]\s*(\d{1,2})\s*[月.\-\/]\s*(\d{1,2})/);
    if (eraMatch) {
      const base = PcaInvoiceWindow.ERA_BASE[eraMatch[1]];
      if (base) {
        return (base + parseInt(eraMatch[2], 10)) + '/' + pad(eraMatch[3]) + '/' + pad(eraMatch[4]);
      }
    }

    const s = raw.replace(/[.\-年月]/g, '/').replace(/日/g, '');

    // 西暦4桁
    const western = s.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
    if (western) return western[1] + '/' + pad(western[2]) + '/' + pad(western[3]);

    // 元号記号なしの1〜2桁年は令和とみなす（手書き請求書の「8年6月30日」など）
    const wareki = s.match(/^\/?(\d{1,2})\/(\d{1,2})\/(\d{1,2})\/?$/);
    if (wareki) {
      const y = parseInt(wareki[1], 10);
      const mo = parseInt(wareki[2], 10);
      const d = parseInt(wareki[3], 10);
      if (y >= 1 && y <= 99 && mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
        return (PcaInvoiceWindow.ERA_BASE['令和'] + y) + '/' + pad(mo) + '/' + pad(d);
      }
    }

    return '';
  }

  /**
   * 蓄積した全明細から最終CSVを組み立てます（複数ファイル→1CSV）。
   * mode: 'internal'（社内財務管理表 取込用） / 'office'（会計事務所 PCA取込用）
   */
  buildCsv(entries, mode, voucherStart) {
    const list = entries || [];
    const isOffice = (mode === 'office');
    const headers = isOffice ? PcaInvoiceWindow.OFFICE_HEADERS : PcaInvoiceWindow.INTERNAL_HEADERS;
    // 伝票番号の開始値。前回出力分の続きから採番するために使います
    // （既に出力済みの伝票番号と重複すると、PCA側の取込で事故になるため）
    let start = Number(voucherStart);
    if (!isFinite(start) || start < 1) start = 1;
    const rows = isOffice ? this.buildOfficeRows(list, start) : this.buildInternalRows(list, start);
    return {
      mode: isOffice ? 'office' : 'internal',
      headers: headers,
      rows: rows,
      csv: CsvUtil.build(headers, rows),
      count: list.length,
      // 科目が未確定の明細（＝あとで設定するもの）を画面に出すための一覧
      unresolved: PcaInvoiceWindow.collectUnresolved(list)
    };
  }

  /**
   * 社内用：アップロードいただいたPCA見本データ（仕訳日記帳）の
   * 借方11列と同じ並びを維持し、末尾に大分類・小分類などを追加しています。
   * そのまま社内財務管理表へ貼り付けられる想定です。
   * 伝票日付は見本にあわせ YYYYMMDD（数値形式）で出力します。
   */
  buildInternalRows(entries, voucherStart) {
    const base = (voucherStart || 1);
    const rows = [];
    entries.forEach(function (e, idx) {
      const voucherNo = base + idx;
      const ymd = e.date ? e.date.replace(/\//g, '') : '';

      // 本体（税抜）行
      rows.push([
        ymd, voucherNo,
        e.accountCode, e.accountName, e.subCode, e.subName,
        e.deptCode, e.deptName, e.taxClass,
        e.netAmount === '' ? '' : e.netAmount,
        e.description, e.major, e.minor, e.vendor, e.sourceFile, e.warnings || ''
      ]);

      // 仮払消費税等 行（税額が読み取れた場合のみ）
      if (e.taxAmount !== '' && e.taxAmount !== 0) {
        rows.push([
          ymd, voucherNo,
          '191', '仮払消費税等', '', '',
          e.deptCode, e.deptName, e.taxClass,
          e.taxAmount,
          e.description, '資産', '仮払消費税等', e.vendor, e.sourceFile, ''
        ]);
      }
    });
    return rows;
  }

  /**
   * 会計事務所用：PCA会計の仕訳インポートを想定した 借方／貸方ペア形式。
   * 貸方は「未払金(420)」を既定の相手科目としています
   * （変更が必要な場合は下記 CREDIT_* の値を書き換えてください）。
   */
  buildOfficeRows(entries, voucherStart) {
    const CREDIT_CODE = '420';
    const CREDIT_NAME = '未払金';
    const CREDIT_TAX = '対象外';
    const base = (voucherStart || 1);
    const rows = [];

    entries.forEach(function (e, idx) {
      const voucherNo = base + idx;
      const memo = [e.vendor, e.description, e.invoiceNo].filter(Boolean).join(' ');

      // 本体（税抜）行
      if (e.netAmount !== '' && e.netAmount !== 0) {
        rows.push([
          e.date, voucherNo,
          e.accountCode, e.accountName, e.subCode, e.subName,
          e.deptCode, e.deptName, e.taxClass, e.netAmount,
          CREDIT_CODE, CREDIT_NAME, '', '',
          e.deptCode, e.deptName, CREDIT_TAX, e.netAmount,
          memo
        ]);
      }

      // 仮払消費税等 行
      if (e.taxAmount !== '' && e.taxAmount !== 0) {
        rows.push([
          e.date, voucherNo,
          '191', '仮払消費税等', '', '',
          e.deptCode, e.deptName, e.taxClass, e.taxAmount,
          CREDIT_CODE, CREDIT_NAME, '', '',
          e.deptCode, e.deptName, CREDIT_TAX, e.taxAmount,
          memo
        ]);
      }
    });
    return rows;
  }
}


// ====================================================================
// SheetExporter：Excel(.xlsx)／Googleスプレッドシート出力の専用クラス
// ------------------------------------------------------------------
// payload = { fileName: 'xxx', sheets: [{ name, headers, rows }] }
// ・toSpreadsheet()   … Googleスプレッドシートを作成しURLを返す
// ・toExcelBase64()   … 一時スプレッドシートをxlsx変換してBase64で返し、
//                       一時ファイルはゴミ箱へ（Drive上に残しません）
// ====================================================================
class SheetExporter {

  _createSpreadsheet(payload) {
    const name = CsvUtil.sanitize(payload && payload.fileName) || 'AI会計コンバーター出力';
    const sheets = (payload && payload.sheets) || [];
    if (sheets.length === 0) throw new Error('出力するデータがありません。');

    const ss = SpreadsheetApp.create(name);

    sheets.forEach(function (def, index) {
      const headers = def.headers || [];
      const rows = def.rows || [];
      const sheetName = CsvUtil.sanitize(def.name) || ('シート' + (index + 1));

      const sheet = (index === 0) ? ss.getSheets()[0] : ss.insertSheet();
      sheet.setName(sheetName.substring(0, 90));

      const table = [];
      if (headers.length) table.push(headers.map(function (h) { return CsvUtil.sanitize(h); }));
      rows.forEach(function (r) {
        const line = [];
        for (let c = 0; c < headers.length; c++) {
          const v = r[c];
          // 数値はそのまま数値として書き込む（Excel側で計算できるように）
          line.push((typeof v === 'number') ? v : CsvUtil.sanitize(v));
        }
        table.push(line);
      });

      if (table.length && headers.length) {
        sheet.getRange(1, 1, table.length, headers.length).setValues(table);
        if (headers.length) {
          sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#eef6fc');
          sheet.setFrozenRows(1);
        }
      }
    });

    SpreadsheetApp.flush();
    return ss;
  }

  toSpreadsheet(payload) {
    const ss = this._createSpreadsheet(payload);
    return { url: ss.getUrl(), id: ss.getId(), name: ss.getName() };
  }

  toExcelBase64(payload) {
    const ss = this._createSpreadsheet(payload);
    const id = ss.getId();
    try {
      const url = 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx';
      const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      if (response.getResponseCode() !== 200) {
        throw new Error('Excel変換に失敗しました (HTTP ' + response.getResponseCode() + ')');
      }
      return {
        base64: Utilities.base64Encode(response.getBlob().getBytes()),
        fileName: ss.getName() + '.xlsx'
      };
    } finally {
      // 一時ファイルは必ず片付ける（失敗しても本処理は止めない）
      try { DriveApp.getFileById(id).setTrashed(true); } catch (cleanupErr) { /* noop */ }
    }
  }
}


// ====================================================================
// 🌐 LogTranslator：画面ログ（エラー）の翻訳・要約専用クラス
// 窓1〜5のいずれの業務ロジックにも依存しない完全に独立したユーティリティ。
// 失敗しても null を返すだけなので、アプリ本体の動作を止めません。
// 翻訳は速度・コストを優先し、常に軽量モデル（Flash-Lite）を使用します。
// ====================================================================
class LogTranslator {
  translate(originalMessage) {
    if (!originalMessage) return null;

    try {
      const client = new GeminiApiClient(ModelRegistry.DEFAULT);
      if (!client.hasApiKey()) return null;

      const prompt = `以下はシステムのログメッセージです。
1. 自然な日本語に翻訳してください（すでに日本語の場合はそのまま使用してください）。
2. その内容を20文字程度で簡潔に要約してください（原因や意味がひと目でわかるように）。
出力は必ず次のJSON形式のみを返してください。挨拶や説明文、Markdownの装飾（\`\`\`など）は一切不要です。
{"translation": "...", "summary": "..."}

対象のログメッセージ：
「${originalMessage}」`;

      const resultText = client.generateText(prompt);
      if (!resultText) return null;

      const jsonMatch = resultText.match(/\{[\s\S]*\}/);
      if (!jsonMatch) return null;

      const parsed = JSON.parse(jsonMatch[0]);
      if (!parsed.translation && !parsed.summary) return null;
      return parsed;
    } catch (e) {
      // 翻訳・要約に失敗してもログ表示自体には影響させない
      return null;
    }
  }
}
