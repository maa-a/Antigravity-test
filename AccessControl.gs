/**
 * AI会計コンバーター ―― メンバー権限と共有作業（AccessControl.gs）
 * ------------------------------------------------------------------
 * ・AccessControl  : 誰が何をしてよいかの判定とメンバー管理
 * ・SharedWorkspace: 窓5の作業内容をメンバー間で共有するための保存庫
 *
 * 【前提】
 * appsscript.json が executeAs: "USER_DEPLOYING" / access: "DOMAIN" の
 * ため、アプリはデプロイした方の権限で動き、同じGoogle Workspace
 * ドメインの方だけがURLを開けます。
 * その中で「誰に何を許すか」をこのファイルで制御します。
 *
 * 【重要】
 * 画面側でボタンを隠すだけでは制限になりません（開発者ツールから
 * 関数を直接呼べてしまうため）。必ずサーバー側のこのクラスで
 * 判定してから処理を行います。
 * ------------------------------------------------------------------
 */

class AccessControl {

  /** 権限の強さ。数字が大きいほど広い権限。 */
  static get RANK() {
    return { viewer: 1, editor: 2, admin: 3 };
  }

  static get ROLE_LABELS() {
    return {
      admin:  '管理者（マスタ編集・メンバー管理・APIキー設定）',
      editor: '編集者（請求書の解析・修正・出力・共有作業の更新）',
      viewer: '閲覧者（共有された内容の閲覧と出力のみ）'
    };
  }

  static get KEY_MEMBERS() { return 'ACL_MEMBERS'; }
  static get KEY_DEFAULT_ROLE() { return 'ACL_DEFAULT_ROLE'; }

  /** アプリをデプロイした人（＝オーナー）。常に管理者として扱います。 */
  static ownerEmail() {
    try {
      return (Session.getEffectiveUser().getEmail() || '').toLowerCase();
    } catch (e) {
      return '';
    }
  }

  /**
   * いまアプリを開いている人のメールアドレス。
   * 同じWorkspaceドメインであれば取得できます。取得できない場合は
   * 空文字を返し、既定ロールが適用されます。
   */
  static currentEmail() {
    try {
      return (Session.getActiveUser().getEmail() || '').toLowerCase();
    } catch (e) {
      return '';
    }
  }

  /** 登録メンバー一覧。 */
  static members() {
    try {
      const raw = PropertiesService.getScriptProperties().getProperty(AccessControl.KEY_MEMBERS);
      if (!raw) return [];
      const list = JSON.parse(raw);
      return Array.isArray(list) ? list : [];
    } catch (e) {
      return [];
    }
  }

  /** メンバー表に載っていない人に適用するロール（既定は閲覧者）。 */
  static defaultRole() {
    const saved = PropertiesService.getScriptProperties().getProperty(AccessControl.KEY_DEFAULT_ROLE);
    return AccessControl.RANK[saved] ? saved : 'viewer';
  }

  /** 指定メールのロールを返す。 */
  static roleOf(email) {
    const target = String(email || '').toLowerCase();
    const owner = AccessControl.ownerEmail();

    // オーナーは必ず管理者（自分自身を締め出して復旧できなくなるのを防ぐ）
    if (target && owner && target === owner) return 'admin';

    const members = AccessControl.members();

    // メンバーが誰も登録されていない間は、全員が編集者として使えるようにする。
    // （導入直後にアプリが使えなくなるのを避けるため）
    if (members.length === 0) return target === owner ? 'admin' : 'editor';

    const hit = members.filter(function (m) {
      return String(m.email || '').toLowerCase() === target;
    })[0];
    if (hit && AccessControl.RANK[hit.role]) return hit.role;

    return AccessControl.defaultRole();
  }

  /** 画面に渡す権限情報一式。 */
  static context() {
    const email = AccessControl.currentEmail();
    const owner = AccessControl.ownerEmail();
    const role = AccessControl.roleOf(email);
    const members = AccessControl.members();

    return {
      email: email,
      ownerEmail: owner,
      isOwner: !!email && email === owner,
      role: role,
      roleLabel: AccessControl.ROLE_LABELS[role] || '',
      canView: AccessControl.allows(role, 'viewer'),
      canEdit: AccessControl.allows(role, 'editor'),
      canAdmin: AccessControl.allows(role, 'admin'),
      memberCount: members.length,
      defaultRole: AccessControl.defaultRole(),
      // メンバー未登録なら「全員が編集者」で動いていることを画面に伝える
      openMode: members.length === 0,
      // メールが取れない環境かどうか（画面に注意書きを出すため）
      identified: !!email
    };
  }

  static allows(role, needed) {
    const have = AccessControl.RANK[role] || 0;
    const want = AccessControl.RANK[needed] || 0;
    return have >= want;
  }

  /**
   * 権限が足りなければ例外を投げる。すべての更新系処理の先頭で呼びます。
   * @param needed 'viewer' | 'editor' | 'admin'
   */
  static require(needed) {
    const ctx = AccessControl.context();
    if (AccessControl.allows(ctx.role, needed)) return ctx;

    const labels = { admin: '管理者', editor: '編集者', viewer: '閲覧者' };
    throw new Error(
      'この操作には「' + (labels[needed] || needed) + '」以上の権限が必要です。' +
      '（現在: ' + (labels[ctx.role] || ctx.role) + (ctx.email ? ' / ' + ctx.email : '') + '）\\n' +
      '権限の変更は管理者にご依頼ください。');
  }

  /** メンバー一覧を保存（管理者のみ）。 */
  static saveMembers(rawList) {
    AccessControl.require('admin');

    const list = [];
    const seen = {};
    const errors = [];

    (rawList || []).forEach(function (m, idx) {
      const email = String((m && m.email) || '').trim().toLowerCase();
      const role = String((m && m.role) || '').trim();
      const note = String((m && m.note) || '').trim();
      if (!email && !role) return;

      if (!email) { errors.push((idx + 1) + '行目: メールアドレスが空です'); return; }
      if (email.indexOf('@') === -1) {
        errors.push((idx + 1) + '行目: メールアドレスの形式が正しくありません（' + email + '）');
        return;
      }
      if (!AccessControl.RANK[role]) {
        errors.push((idx + 1) + '行目: 権限を選んでください（' + email + '）');
        return;
      }
      if (seen[email]) return;   // 重複は1件にまとめる
      seen[email] = true;
      list.push({ email: email, role: role, note: note });
    });

    if (errors.length) {
      throw new Error('保存できませんでした。\\n・' + errors.slice(0, 10).join('\\n・'));
    }

    // 管理者が1人もいなくなると誰も設定を戻せなくなるため、
    // オーナーは常に管理者として自動的に補います。
    const owner = AccessControl.ownerEmail();
    const hasAdmin = list.some(function (m) { return m.role === 'admin'; });
    if (owner && !list.some(function (m) { return m.email === owner; })) {
      list.unshift({ email: owner, role: 'admin', note: 'アプリのオーナー（自動登録）' });
    } else if (!hasAdmin && owner) {
      list.forEach(function (m) { if (m.email === owner) m.role = 'admin'; });
    }

    PropertiesService.getScriptProperties()
      .setProperty(AccessControl.KEY_MEMBERS, JSON.stringify(list));

    return { count: list.length, context: AccessControl.context(), members: list };
  }

  /** メンバー表に無い人の既定ロールを保存（管理者のみ）。 */
  static setDefaultRole(role) {
    AccessControl.require('admin');
    if (!AccessControl.RANK[role]) throw new Error('権限の指定が正しくありません: ' + role);
    PropertiesService.getScriptProperties().setProperty(AccessControl.KEY_DEFAULT_ROLE, role);
    return { defaultRole: role, context: AccessControl.context() };
  }
}


// ====================================================================
// SharedWorkspace：窓5の作業内容をメンバー間で共有するための保存庫
// ------------------------------------------------------------------
// ・保存先は Drive 上のJSONファイル1つ（オーナーのDrive内）。
//   ScriptProperties は1値9KBの制限があり、明細を保存しきれないためです。
// ・同時編集で内容が壊れないよう、保存のたびに版番号(revision)を上げ、
//   古い版に基づく上書きは拒否します。
// ・排他制御に LockService を使い、同時保存による破損を防ぎます。
// ====================================================================
class SharedWorkspace {

  static get KEY_FILE_ID() { return 'SHARED_WORK_FILE_ID'; }
  static get FILE_NAME() { return 'AI会計コンバーター_共有作業データ.json'; }

  static _emptyState() {
    return { revision: 0, savedAt: '', savedBy: '', data: null };
  }

  /** 保存用ファイルを取得（無ければ作成）。 */
  static _file() {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty(SharedWorkspace.KEY_FILE_ID);
    if (id) {
      try {
        const file = DriveApp.getFileById(id);
        if (!file.isTrashed()) return file;
      } catch (e) { /* 消された場合は作り直す */ }
    }
    const created = DriveApp.createFile(
      SharedWorkspace.FILE_NAME,
      JSON.stringify(SharedWorkspace._emptyState()),
      MimeType.PLAIN_TEXT);
    props.setProperty(SharedWorkspace.KEY_FILE_ID, created.getId());
    return created;
  }

  /** 共有作業を読み込む。閲覧者以上。 */
  static load() {
    AccessControl.require('viewer');
    try {
      const text = SharedWorkspace._file().getBlob().getDataAsString();
      const state = JSON.parse(text);
      if (!state || typeof state.revision !== 'number') return SharedWorkspace._emptyState();
      return state;
    } catch (e) {
      return SharedWorkspace._emptyState();
    }
  }

  /**
   * 共有作業を保存する。編集者以上。
   * @param expectedRevision 画面が読み込んだときの版番号。
   *        これが現在の版と違う場合は、他の人が更新したとみなして拒否します
   *        （force が true のときのみ上書きします）。
   */
  static save(data, expectedRevision, force) {
    const ctx = AccessControl.require('editor');

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) {
      throw new Error('他の人が保存中です。少し待ってからもう一度お試しください。');
    }
    try {
      const file = SharedWorkspace._file();
      let current;
      try {
        current = JSON.parse(file.getBlob().getDataAsString());
      } catch (e) {
        current = SharedWorkspace._emptyState();
      }
      const currentRev = (current && typeof current.revision === 'number') ? current.revision : 0;

      if (!force && typeof expectedRevision === 'number' && expectedRevision !== currentRev) {
        throw new Error(
          '他の人が先に共有データを更新しています（' +
          (current.savedBy || '不明') + ' / ' + (current.savedAt || '') + '）。\\n' +
          '「共有から読み込む」で最新を取り込んでから、もう一度保存してください。');
      }

      const next = {
        revision: currentRev + 1,
        savedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
        savedBy: ctx.email || '（不明）',
        data: data || null
      };
      file.setContent(JSON.stringify(next));
      return { revision: next.revision, savedAt: next.savedAt, savedBy: next.savedBy };
    } finally {
      lock.releaseLock();
    }
  }

  /** 共有作業を空にする。編集者以上。 */
  static clear() {
    const ctx = AccessControl.require('editor');
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) {
      throw new Error('他の人が保存中です。少し待ってからもう一度お試しください。');
    }
    try {
      const file = SharedWorkspace._file();
      let currentRev = 0;
      try {
        currentRev = JSON.parse(file.getBlob().getDataAsString()).revision || 0;
      } catch (e) { /* 読めなければ0から */ }

      const next = {
        revision: currentRev + 1,
        savedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
        savedBy: ctx.email || '（不明）',
        data: null
      };
      file.setContent(JSON.stringify(next));
      return { revision: next.revision, savedAt: next.savedAt, savedBy: next.savedBy };
    } finally {
      lock.releaseLock();
    }
  }
}
