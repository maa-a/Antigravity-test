/**
 * AI会計コンバーター V8.00 ―― マスタ保存・科目解決（MasterStore.gs）
 * ------------------------------------------------------------------
 * ・MasterStore     : マスタの保存／読込／Excel貼り付け取り込み
 * ・AccountResolver : AIが返した科目名をマスタ上の正規科目へ突き合わせる
 *
 * 【最重要仕様】
 * マスタに存在しない・判定できない科目は【必ず空欄】を返します。
 * 推測での当てはめや、化けた文字列をそのまま出力することはしません
 * （後から人が書き換える前提のため）。
 *
 * このファイルは窓1〜窓4（通帳・クレカ・弥生変換）の処理には一切関与しません。
 * ------------------------------------------------------------------
 */

/**
 * このファイル内だけで使う軽量ヘルパー。
 * マスタ編集は業務ロジック（窓1〜5）から独立させたいので、
 * Code.gs の CsvUtil には依存させずここに最小限だけ持ちます。
 */
class CsvUtilLite {
  /** 値を文字列にして前後の空白と制御文字を落とす。null/undefined は ''。 */
  static trim(value) {
    if (value === null || value === undefined) return '';
    return String(value)
      .replace(/[\u0000-\u001F\u007F]/g, '')
      .replace(/\u3000/g, ' ')
      .trim();
  }
}


class MasterStore {

  static get KEYS() {
    return {
      pcaAccounts:      'MASTER_PCA_ACCOUNTS',
      departments:      'MASTER_DEPARTMENTS',
      taxClasses:       'MASTER_TAX_CLASSES',
      subAccounts:      'MASTER_SUB_ACCOUNTS',
      internalAccounts: 'MASTER_INTERNAL_ACCOUNTS',
      keywordRules:     'MASTER_KEYWORD_RULES'
    };
  }

  static get SEEDS() {
    return {
      pcaAccounts:      MasterSeed.PCA_ACCOUNTS,
      departments:      MasterSeed.DEPARTMENTS,
      taxClasses:       MasterSeed.TAX_CLASSES,
      subAccounts:      MasterSeed.SUB_ACCOUNTS,
      internalAccounts: MasterSeed.INTERNAL_ACCOUNTS,
      keywordRules:     MasterSeed.KEYWORD_RULES
    };
  }

  /**
   * 1マスタを読み込む。保存済みが無い／壊れている場合はシード値を返す。
   * ここで例外を投げないことで、マスタ破損時もアプリが止まらないようにしています。
   */
  static load(name) {
    const key = MasterStore.KEYS[name];
    if (!key) return null;
    try {
      const raw = PropertiesService.getScriptProperties().getProperty(key);
      if (!raw) return MasterStore.SEEDS[name];
      const parsed = JSON.parse(raw);
      if (parsed === null || parsed === undefined) return MasterStore.SEEDS[name];
      if (Array.isArray(parsed) && parsed.length === 0) return MasterStore.SEEDS[name];
      return parsed;
    } catch (e) {
      return MasterStore.SEEDS[name];
    }
  }

  static save(name, value) {
    const key = MasterStore.KEYS[name];
    if (!key) throw new Error('未知のマスタ名です: ' + name);
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(value));
  }

  static reset(name) {
    const key = MasterStore.KEYS[name];
    if (!key) throw new Error('未知のマスタ名です: ' + name);
    PropertiesService.getScriptProperties().deleteProperty(key);
  }

  /** 画面（窓5）が必要とするマスタを一括で返す。 */
  static loadAll() {
    return {
      pcaAccounts:      MasterStore.load('pcaAccounts'),
      departments:      MasterStore.load('departments'),
      taxClasses:       MasterStore.load('taxClasses'),
      subAccounts:      MasterStore.load('subAccounts'),
      internalAccounts: MasterStore.load('internalAccounts'),
      keywordRules:     MasterStore.load('keywordRules')
    };
  }

  /**
   * Excel／スプレッドシートからコピーしたテキスト（タブ区切り or カンマ区切り）を
   * 行配列に変換する共通パーサ。空行と見出し行らしき行は自動で除外します。
   */
  static parsePastedTable(text) {
    if (!text) return [];
    const rows = [];
    String(text).split(/\r\n|\n|\r/).forEach(function (line) {
      if (!line || !line.trim()) return;
      const hasTab = line.indexOf('\t') !== -1;
      const cols = (hasTab ? line.split('\t') : line.split(','))
        .map(function (c) { return String(c).replace(/^"|"$/g, '').trim(); });
      if (cols.every(function (c) { return c === ''; })) return;
      rows.push(cols);
    });
    if (rows.length === 0) return [];
    const head = rows[0].join('');
    if (/大分類|小分類|科目|コード|部門|税区分|補助/.test(head) && !/^[0-9]/.test(rows[0][0] || '')) {
      rows.shift();
    }
    return rows;
  }

  /**
   * 社内管理科目（社内財務管理表「PL（詳細）」D列＝大分類／E列＝小分類）の取り込み。
   * 期待する列：[大分類, 小分類, (任意)PCA勘定科目コード, (任意)セクション]
   *
   * 「PL（詳細）」のD列は結合セル（グループ見出し行にだけ値が入り、
   * 明細行は空欄）になっているため、D列が空の行は直前のD列の値を
   * 引き継ぎます（Excelから貼り付けたそのままの形で取り込めます）。
   * E列が空の行は、D列の値をそのまま小分類として扱います。
   */
  static importInternalAccounts(text) {
    const rows = MasterStore.parsePastedTable(text);
    const list = [];
    const seen = {};
    let lastMajor = '';

    rows.forEach(function (cols) {
      const major = (cols[0] || '').trim() || lastMajor;
      if (!major) return;
      lastMajor = major;

      const minor = (cols[1] || '').trim() || major;
      const dedupeKey = major + '||' + minor;
      if (seen[dedupeKey]) return;
      seen[dedupeKey] = true;

      list.push({
        section: (cols[3] || '').trim(),
        major: major,
        minor: minor,
        pcaCode: (cols[2] || '').trim()
      });
    });

    if (list.length === 0) {
      throw new Error('取り込める行がありませんでした。D列（大分類）とE列（小分類）を含めて貼り付けてください。');
    }
    MasterStore.save('internalAccounts', list);
    return {
      count: list.length,
      unlinked: list.filter(function (x) { return !x.pcaCode; }).length
    };
  }

  /**
   * PCA勘定科目マスタの取り込み。期待する列：[コード, 科目名, (任意)大分類]
   */
  static importPcaAccounts(text) {
    const rows = MasterStore.parsePastedTable(text);
    const list = [];
    rows.forEach(function (cols) {
      const code = (cols[0] || '').trim();
      const name = (cols[1] || '').trim();
      if (!code || !name) return;
      list.push({ code: code, name: name, major: (cols[2] || '').trim() });
    });
    if (list.length === 0) {
      throw new Error('取り込める行がありませんでした。「コード」「科目名」の2列以上を貼り付けてください。');
    }
    MasterStore.save('pcaAccounts', list);
    return { count: list.length };
  }

  // ==================================================================
  // ここから：画面の表UIから1行ずつ編集・追加・削除するための保存処理
  // ------------------------------------------------------------------
  // 「最初は空欄で出しておいて、決まったものから足していく」運用のため、
  // 全置換の貼り付け取り込みとは別に、編集後の一覧をそのまま保存する
  // 経路を用意しています。保存前に必ず検証し、問題があれば保存せずに
  // 理由を返します（壊れたマスタが保存されて業務が止まらないように）。
  // ==================================================================

  /** マスタごとの重複判定キー。 */
  static keyOf(name, item) {
    if (name === 'internalAccounts') {
      return AccountResolver.normalize(item.major) + '||' + AccountResolver.normalize(item.minor);
    }
    if (name === 'pcaAccounts' || name === 'departments') {
      return String(item.code === undefined ? '' : item.code).trim();
    }
    if (name === 'keywordRules') {
      // 取引先＋摘要の組み合わせで1つのルールとして扱う
      return AccountResolver.normalize(item.vendorKeyword) + '||' + AccountResolver.normalize(item.keyword);
    }
    return JSON.stringify(item);
  }

  /**
   * 画面から渡された一覧を検証して正規化する。
   * 戻り値：{ list: 正規化済み一覧, errors: [エラー文言], warnings: [注意文言] }
   */
  static validate(name, rawList) {
    const list = [];
    const errors = [];
    const warnings = [];
    const seen = {};
    const input = rawList || [];

    // PCA科目コードの妥当性チェック用（社内管理科目・仕分けルールで使用）
    const pcaCodes = {};
    (MasterStore.load('pcaAccounts') || []).forEach(function (a) {
      pcaCodes[String(a.code).trim()] = true;
    });

    input.forEach(function (raw, idx) {
      const lineNo = idx + 1;
      const item = {};

      if (name === 'internalAccounts') {
        item.section = CsvUtilLite.trim(raw.section);
        item.major = CsvUtilLite.trim(raw.major);
        item.minor = CsvUtilLite.trim(raw.minor);
        item.pcaCode = CsvUtilLite.trim(raw.pcaCode);
        if (!item.major && !item.minor) return;          // 空行は無視
        if (!item.major) { errors.push(lineNo + '行目: 大分類が空です'); return; }
        if (!item.minor) item.minor = item.major;        // 販管費のように小分類＝大分類のケース
        if (item.pcaCode && !pcaCodes[item.pcaCode]) {
          errors.push(lineNo + '行目: PCA勘定科目コード「' + item.pcaCode + '」はPCA勘定科目マスタにありません');
          return;
        }
        if (!item.pcaCode) {
          warnings.push('「' + item.major + ' / ' + item.minor + '」はPCA科目が未設定です（会計事務所用CSVでは空欄で出力されます）');
        }

      } else if (name === 'pcaAccounts') {
        item.code = CsvUtilLite.trim(raw.code);
        item.name = CsvUtilLite.trim(raw.name);
        item.major = CsvUtilLite.trim(raw.major);
        if (!item.code && !item.name) return;
        if (!item.code) { errors.push(lineNo + '行目: 勘定科目コードが空です'); return; }
        if (!item.name) { errors.push(lineNo + '行目: 勘定科目名が空です'); return; }

      } else if (name === 'departments') {
        item.code = CsvUtilLite.trim(raw.code);
        item.name = CsvUtilLite.trim(raw.name);
        if (item.code === '' && item.name === '') return;
        if (!item.name) { errors.push(lineNo + '行目: 部門名が空です'); return; }

      } else if (name === 'keywordRules') {
        item.vendorKeyword = CsvUtilLite.trim(raw.vendorKeyword);
        item.keyword = CsvUtilLite.trim(raw.keyword);
        item.major = CsvUtilLite.trim(raw.major);
        item.minor = CsvUtilLite.trim(raw.minor);
        item.note = CsvUtilLite.trim(raw.note);
        if (!item.vendorKeyword && !item.keyword && !item.major && !item.minor) return;
        if (!item.vendorKeyword && !item.keyword) {
          errors.push(lineNo + '行目: 取引先キーワードか摘要キーワードのどちらかは入力してください');
          return;
        }
        const label = [item.vendorKeyword, item.keyword].filter(Boolean).join(' + ');
        if (!item.major || !item.minor) {
          errors.push(lineNo + '行目: ルール「' + label + '」の大分類・小分類を選んでください');
          return;
        }

      } else {
        errors.push('未知のマスタ名です: ' + name);
        return;
      }

      const key = MasterStore.keyOf(name, item);
      if (seen[key]) {
        warnings.push(lineNo + '行目: 重複していたため1件にまとめました');
        return;
      }
      seen[key] = true;
      list.push(item);
    });

    // 仕分けルールの参照先が社内管理科目に存在するかを確認
    if (name === 'keywordRules') {
      const pairs = {};
      (MasterStore.load('internalAccounts') || []).forEach(function (a) {
        pairs[AccountResolver.normalize(a.major) + '||' + AccountResolver.normalize(a.minor)] = true;
      });
      list.forEach(function (rule) {
        const k = AccountResolver.normalize(rule.major) + '||' + AccountResolver.normalize(rule.minor);
        if (!pairs[k]) {
          const label = [rule.vendorKeyword, rule.keyword].filter(Boolean).join(' + ');
          errors.push('ルール「' + label + '」の科目「' + rule.major + ' / ' + rule.minor + '」が社内管理科目にありません');
        }
      });
    }

    return { list: list, errors: errors, warnings: warnings };
  }

  /**
   * 画面の表UIで編集した一覧を保存する。
   * 検証エラーがある場合は保存せず例外を投げます（部分保存で
   * マスタが壊れるのを防ぐため）。
   */
  static saveMasterList(name, rawList) {
    if (!MasterStore.KEYS[name]) throw new Error('未知のマスタ名です: ' + name);
    const result = MasterStore.validate(name, rawList);
    if (result.errors.length) {
      throw new Error('保存できませんでした。\n・' + result.errors.slice(0, 10).join('\n・'));
    }
    if (result.list.length === 0) {
      throw new Error('保存できる行がありません。1行以上入力してください。');
    }
    MasterStore.save(name, result.list);
    return {
      count: result.list.length,
      warnings: result.warnings,
      masters: MasterStore.loadAll()
    };
  }

  /**
   * 仕分けルールを1件だけ追加（同じキーワードがあれば上書き）。
   * 窓5の「未確定リスト」から、その場でルール化するために使います。
   */
  static addKeywordRule(rule) {
    const list = (MasterStore.load('keywordRules') || []).slice();
    const incoming = {
      vendorKeyword: CsvUtilLite.trim(rule && rule.vendorKeyword),
      keyword: CsvUtilLite.trim(rule && rule.keyword),
      major: CsvUtilLite.trim(rule && rule.major),
      minor: CsvUtilLite.trim(rule && rule.minor),
      note: CsvUtilLite.trim(rule && rule.note)
    };
    if (!incoming.vendorKeyword && !incoming.keyword) {
      throw new Error('取引先キーワードか摘要キーワードのどちらかは入力してください。');
    }
    if (!incoming.major || !incoming.minor) throw new Error('大分類と小分類を選んでください。');

    const key = MasterStore.keyOf('keywordRules', incoming);
    const next = list.filter(function (x) {
      return MasterStore.keyOf('keywordRules', x) !== key;
    });
    next.push(incoming);

    const result = MasterStore.validate('keywordRules', next);
    if (result.errors.length) throw new Error(result.errors[0]);

    MasterStore.save('keywordRules', result.list);
    return { count: result.list.length, masters: MasterStore.loadAll() };
  }

  /**
   * 社内管理科目を1件だけ追加・更新（同じ大分類＋小分類なら上書き）。
   * 未確定リストから「この科目を新しく作る」ときに使います。
   */
  static addInternalAccount(item) {
    const list = (MasterStore.load('internalAccounts') || []).slice();
    const incoming = {
      section: CsvUtilLite.trim(item && item.section),
      major: CsvUtilLite.trim(item && item.major),
      minor: CsvUtilLite.trim(item && item.minor),
      pcaCode: CsvUtilLite.trim(item && item.pcaCode)
    };
    if (!incoming.major) throw new Error('大分類を入力してください。');
    if (!incoming.minor) incoming.minor = incoming.major;

    const key = MasterStore.keyOf('internalAccounts', incoming);
    const next = list.filter(function (x) {
      return MasterStore.keyOf('internalAccounts', x) !== key;
    });
    next.push(incoming);

    const result = MasterStore.validate('internalAccounts', next);
    if (result.errors.length) throw new Error(result.errors[0]);

    MasterStore.save('internalAccounts', result.list);
    return { count: result.list.length, masters: MasterStore.loadAll() };
  }

  /** 部門マスタの取り込み。期待する列：[部門コード, 部門名] */
  static importDepartments(text) {
    const rows = MasterStore.parsePastedTable(text);
    const list = [];
    rows.forEach(function (cols) {
      const code = (cols[0] || '').trim();
      const name = (cols[1] || '').trim();
      if (code === '' && name === '') return;
      list.push({ code: code, name: name });
    });
    if (list.length === 0) {
      throw new Error('取り込める行がありませんでした。「部門コード」「部門名」の2列を貼り付けてください。');
    }
    MasterStore.save('departments', list);
    return { count: list.length };
  }
}


// ====================================================================
// AccountResolver：AIの返した科目名をマスタ上の正規科目へ突き合わせる。
// 該当なしは常に null（＝出力は空欄）。推測での当てはめは行いません。
// ====================================================================
class AccountResolver {

  constructor(masters) {
    this.masters = masters || MasterStore.loadAll();

    this._pcaByName = {};
    this._pcaByCode = {};
    (this.masters.pcaAccounts || []).forEach(function (a) {
      this._pcaByName[AccountResolver.normalize(a.name)] = a;
      this._pcaByCode[String(a.code).trim()] = a;
    }, this);

    // 科目の識別キーは【大分類＋小分類のペア】。
    // 「版権元RY」のように、同じ小分類名が複数の大分類に存在するためです。
    this._internalByPair = {};
    // 小分類名が全科目を通じて一意なものだけ、名前だけでも引けるようにする。
    // 重複する小分類名は、大分類が分からないと確定できないので登録しません
    // （＝当てずっぽうで拾わない＝空欄になる）。
    const minorCount = {};
    (this.masters.internalAccounts || []).forEach(function (a) {
      if (!a.minor) return;
      const k = AccountResolver.normalize(a.minor);
      minorCount[k] = (minorCount[k] || 0) + 1;
    });

    this._internalByMinor = {};
    this._ambiguousMinors = {};
    (this.masters.internalAccounts || []).forEach(function (a) {
      if (!a.minor) return;
      const minorKey = AccountResolver.normalize(a.minor);
      this._internalByPair[AccountResolver.normalize(a.major) + '||' + minorKey] = a;
      if (minorCount[minorKey] === 1) {
        this._internalByMinor[minorKey] = a;
      } else {
        this._ambiguousMinors[minorKey] = true;
      }
    }, this);

    this._deptByName = {};
    (this.masters.departments || []).forEach(function (d) {
      this._deptByName[AccountResolver.normalize(d.name)] = d;
    }, this);

    this._taxSet = {};
    (this.masters.taxClasses || []).forEach(function (t) {
      this._taxSet[AccountResolver.normalize(t)] = t;
    }, this);
  }

  /**
   * 表記ゆれ吸収用の正規化。
   * 全角英数・カッコ・ダッシュ・空白のゆれを潰したうえで比較します。
   * （科目名が一致せず空欄だらけになるのを防ぐための処理）
   */
  static normalize(value) {
    if (value === null || value === undefined) return '';
    let s = String(value);
    s = s.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (ch) {
      return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
    });
    s = s.replace(/（/g, '(').replace(/）/g, ')');
    s = s.replace(/[〜～~]/g, '~');
    s = s.replace(/[－ー−―‐—-]/g, '-');
    s = s.replace(/[\s　]/g, '');
    return s.toLowerCase();
  }

  /** PCA勘定科目を名前またはコードから解決。見つからなければ null。 */
  resolvePcaAccount(nameOrCode) {
    if (!nameOrCode) return null;
    const raw = String(nameOrCode).trim();
    if (this._pcaByCode[raw]) return this._pcaByCode[raw];
    const key = AccountResolver.normalize(raw);
    if (this._pcaByName[key]) return this._pcaByName[key];
    return null;
  }

  /**
   * 社内管理科目を解決。
   * 1) 大分類＋小分類のペアで一致 → 確定
   * 2) 小分類だけで一致し、その小分類名が全科目で一意 → 確定
   * 3) 小分類名が複数の大分類に存在する（例：版権元RY）のに大分類が
   *    分からない → null（＝空欄）。当てずっぽうで片方を選びません。
   */
  resolveInternal(minorName, majorName) {
    if (!minorName) return null;
    const minorKey = AccountResolver.normalize(minorName);

    if (majorName) {
      const pair = this._internalByPair[AccountResolver.normalize(majorName) + '||' + minorKey];
      if (pair) return pair;
    }
    return this._internalByMinor[minorKey] || null;
  }

  /** その小分類名が複数の大分類に存在する（＝大分類なしでは確定できない）か。 */
  isAmbiguousMinor(minorName) {
    if (!minorName) return false;
    return !!this._ambiguousMinors[AccountResolver.normalize(minorName)];
  }

  /** 部門を解決。見つからなければ null。 */
  resolveDepartment(nameOrCode) {
    if (nameOrCode === null || nameOrCode === undefined || nameOrCode === '') return null;
    const raw = String(nameOrCode).trim();
    const byName = this._deptByName[AccountResolver.normalize(raw)];
    if (byName) return byName;
    const byCode = (this.masters.departments || []).filter(function (d) {
      return String(d.code).trim() === raw;
    })[0];
    return byCode || null;
  }

  /** 税区分を解決。マスタに無い文字列は null（＝空欄）。 */
  resolveTax(taxName) {
    if (!taxName) return null;
    return this._taxSet[AccountResolver.normalize(taxName)] || null;
  }

  /** 補助科目を解決。見つからなければ null。 */
  resolveSubAccount(pcaCode, subName) {
    if (!pcaCode || !subName) return null;
    const list = (this.masters.subAccounts || {})[String(pcaCode).trim()];
    if (!list) return null;
    const key = AccountResolver.normalize(subName);
    const hit = list.filter(function (pair) {
      return AccountResolver.normalize(pair[1]) === key;
    })[0];
    return hit ? { code: hit[0], name: hit[1] } : null;
  }

  /**
   * 取引先名と摘要から自動仕分けルールを当て、{ major, minor, rule } を返す。
   * 該当なしなら null。AIの判断より優先して適用します（社内ルールの固定化）。
   *
   * ・vendorKeyword … 取引先名に含まれる文字（空なら取引先は問わない）
   * ・keyword       … 摘要・品名に含まれる文字（空なら摘要は問わない）
   * ・両方指定した場合は【両方に一致】したときだけ適用（安全側）
   *
   * ＜適用の優先順位＞
   * もっとも具体的なルールが勝ちます。
   *   1. 取引先＋摘要の両方を指定したルール
   *   2. 取引先だけのルール
   *   3. 摘要だけのルール
   * 同じ具体度なら、条件の文字数が長い（＝限定的な）ルールが優先されます。
   * これにより「北日本カコー かつ 前受金 → ◯◯」のような個別ルールが、
   * 「前受金 → △△」のような一般ルールより先に適用されます。
   */
  applyKeywordRule(description, vendor) {
    const descKey = AccountResolver.normalize(description);
    const vendorKey = AccountResolver.normalize(vendor);
    // 後方互換：取引先が渡されない呼び出しでは、摘要側で両方を見る
    const anyKey = vendorKey ? (descKey + ' ' + vendorKey) : descKey;
    if (!descKey && !vendorKey) return null;

    const rules = (this.masters.keywordRules || []).map(function (rule, index) {
      const v = AccountResolver.normalize(rule.vendorKeyword);
      const k = AccountResolver.normalize(rule.keyword);
      return {
        rule: rule,
        index: index,
        vendorKey: v,
        keyKey: k,
        // 具体度：両方指定=2 / 取引先のみ=1 / 摘要のみ=0
        specificity: (v && k) ? 2 : (v ? 1 : 0),
        length: v.length + k.length
      };
    }).filter(function (r) {
      return r.vendorKey || r.keyKey;
    });

    rules.sort(function (a, b) {
      if (b.specificity !== a.specificity) return b.specificity - a.specificity;
      if (b.length !== a.length) return b.length - a.length;
      return a.index - b.index;   // 同条件なら登録順で安定させる
    });

    for (let i = 0; i < rules.length; i++) {
      const r = rules[i];
      if (r.vendorKey && vendorKey.indexOf(r.vendorKey) === -1) continue;
      // 取引先条件つきのルールは、摘要を摘要欄だけで判定する（誤爆を防ぐ）
      const haystack = r.vendorKey ? descKey : anyKey;
      if (r.keyKey && haystack.indexOf(r.keyKey) === -1) continue;
      return {
        major: r.rule.major || '',
        minor: r.rule.minor || '',
        rule: r.rule
      };
    }
    return null;
  }
}
