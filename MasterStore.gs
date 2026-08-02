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
   * 期待する列：[大分類, 小分類, (任意)PCA勘定科目コード]
   */
  static importInternalAccounts(text) {
    const rows = MasterStore.parsePastedTable(text);
    const list = [];
    const seen = {};
    rows.forEach(function (cols) {
      const major = cols[0] || '';
      const minor = cols[1] || '';
      if (!major && !minor) return;
      const dedupeKey = major + '||' + minor;
      if (seen[dedupeKey]) return;
      seen[dedupeKey] = true;
      list.push({ major: major, minor: minor, pcaCode: (cols[2] || '').trim() });
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

    this._internalByMinor = {};
    (this.masters.internalAccounts || []).forEach(function (a) {
      if (a.minor) this._internalByMinor[AccountResolver.normalize(a.minor)] = a;
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

  /** 社内管理科目（小分類）から解決。見つからなければ null。 */
  resolveInternal(minorName) {
    if (!minorName) return null;
    return this._internalByMinor[AccountResolver.normalize(minorName)] || null;
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
   * 摘要文からキーワードルールで小分類を推定。該当なしなら null。
   * AIの判断より優先して適用します（社内ルールを固定化するため）。
   */
  applyKeywordRule(description) {
    if (!description) return null;
    const target = AccountResolver.normalize(description);
    const rules = this.masters.keywordRules || [];
    for (let i = 0; i < rules.length; i++) {
      const kw = AccountResolver.normalize(rules[i].keyword);
      if (kw && target.indexOf(kw) !== -1) return rules[i].minor;
    }
    return null;
  }
}
