/**
 * Google Drive 共有ドライブ情報収集ツール
 *
 * 機能:
 * - 全共有ドライブの情報を取得
 * - 各共有ドライブのフォルダ・ファイル構造を再帰的に取得
 * - 権限情報と外部共有状況を調査
 * - 結果をスプレッドシートに出力
 *
 * 作成者: GIC技術部
 * 作成日: 2024年9月26日
 * バージョン: 1.3
 *
 * 変更履歴:
 * v1.3 (2025/09/28): 権限表示の大幅改善
 *   - 「管理者」列を追加（コンテンツ管理者の左）
 *   - フォルダ: 上位権限+個別権限を組み合わせて表示
 *   - ファイル: 管理者/コンテンツ管理者/投稿者に上位権限を表示
 *   - 重複除去機能を実装（combinePermissions関数）
 * v1.2 (2025/09/30): useDomainAdminAccess オプション追加
 * v1.1 (2025/09/30): 権限取得機能強化
 * v1.0 (2025/09/29): 初回リリース
 */

// =============================================================================
// 設定
// =============================================================================

const CONFIG = {
  // スプレッドシート設定
  SPREADSHEET_NAME: 'Drive情報収集結果',

  // API制限対策
  BATCH_SIZE: 1000,           // 一度にスプレッドシートに書き込む件数
  API_DELAY: 100,             // API呼び出し間の待機時間(ms)
  MAX_EXECUTION_TIME: 330,    // 最大実行時間(秒) 6分制限対策

  // フォルダ階層設定
  MAX_DEPTH: 10,              // 最大階層深度

  // 権限設定
  ALLOWED_USERS: [
    '*****@*****.co.jp'   // 実行許可ユーザーのメールアドレス
  ],

  // 会社ドメイン（外部共有判定用）
  COMPANY_DOMAINS: ['*****.co.jp']
};

// =============================================================================
// メイン実行関数
// =============================================================================

/**
 * メイン実行関数
 * スプレッドシートのボタンから呼び出される
 */
function main() {
  try {
    console.log('=== Drive情報収集ツール開始 ===');

    // 権限チェック
    if (!checkExecutionPermission()) {
      throw new Error('実行権限がありません。管理者にお問い合わせください。');
    }

    const startTime = new Date();
    const spreadsheet = getOrCreateSpreadsheet();

    // スプレッドシートIDを保存
    PropertiesService.getScriptProperties().setProperty('CURRENT_SPREADSHEET_ID', spreadsheet.getId());

    // 実行状況を初期化
    updateExecutionStatus('実行開始', 0);

    // フェーズ1: 共有ドライブ一覧取得
    updateExecutionStatus('共有ドライブ一覧取得中...', 10);
    const sharedDrives = getSharedDrives();
    console.log(`共有ドライブ数: ${sharedDrives.length}`);

    // マスターシート作成・更新
    updateExecutionStatus('マスターシート作成中...', 20);
    createMasterSheet(spreadsheet, sharedDrives);

    // フェーズ2: 各共有ドライブの詳細情報取得
    const driveStats = {};
    for (let i = 0; i < sharedDrives.length; i++) {
      const drive = sharedDrives[i];
      const progress = 20 + (i / sharedDrives.length) * 70;

      updateExecutionStatus(`${drive.name} 処理中... (${i + 1}/${sharedDrives.length})`, progress);

      try {
        // 実行時間チェック
        if (isTimeoutApproaching(startTime)) {
          console.log('実行時間制限に近づいています。処理を一時停止します。');
          saveProgress(i, sharedDrives);
          updateExecutionStatus('一時停止 - 続行ボタンで再開してください', progress);
          return;
        }

        // ドライブ詳細情報取得
        const stats = processSingleDrive(spreadsheet, drive, i + 1);
        driveStats[drive.id] = stats;

      } catch (error) {
        console.error(`ドライブ ${drive.name} の処理でエラー:`, error);
        logError(drive.name, error.toString());
      }
    }

    // マスターシートを統計情報で更新
    updateMasterSheetWithStats(spreadsheet, sharedDrives, driveStats);

    // 完了処理
    const endTime = new Date();
    const executionTime = Math.round((endTime - startTime) / 1000);

    updateExecutionStatus(`完了 (実行時間: ${executionTime}秒)`, 100);
    console.log('=== Drive情報収集ツール完了 ===');

  } catch (error) {
    console.error('メイン処理でエラー:', error);
    updateExecutionStatus(`エラー: ${error.message}`, 0);
    throw error;
  }
}

/**
 * 続行実行関数
 * 一時停止した処理を再開
 */
function continueExecution() {
  try {
    const progress = getProgressFromProperties();
    if (!progress) {
      throw new Error('続行可能な処理が見つかりません。');
    }

    console.log('処理を再開します...');
    // 実装: 保存された進捗から処理再開

  } catch (error) {
    console.error('続行処理でエラー:', error);
    updateExecutionStatus(`続行エラー: ${error.message}`, 0);
  }
}

// =============================================================================
// 共有ドライブ取得
// =============================================================================

/**
 * 全共有ドライブを取得
 * @returns {Array} 共有ドライブのリスト
 */
function getSharedDrives() {
  try {
    const drives = [];
    let pageToken = null;

    do {
      const params = {
        pageSize: 100,
        fields: 'nextPageToken,drives(id,name,createdTime,capabilities)',
        useDomainAdminAccess: true  // ドメイン管理者として全共有ドライブを取得
      };

      if (pageToken) {
        params.pageToken = pageToken;
      }

      const response = Drive.Drives.list(params);

      if (response.drives) {
        drives.push(...response.drives);
      }

      pageToken = response.nextPageToken;

    } while (pageToken);

    return drives;

  } catch (error) {
    console.error('共有ドライブ取得エラー:', error);
    throw new Error(`共有ドライブの取得に失敗しました: ${error.message}`);
  }
}

// =============================================================================
// 個別ドライブ処理
// =============================================================================

/**
 * 単一の共有ドライブを処理
 * @param {Spreadsheet} spreadsheet スプレッドシート
 * @param {Object} drive ドライブ情報
 * @param {number} index ドライブのインデックス
 */
function processSingleDrive(spreadsheet, drive, index) {
  console.log(`処理開始: ${drive.name} (ID: ${drive.id})`);

  // ドライブ専用シート作成
  const sheetName = `${String(index).padStart(2, '0')}_${sanitizeSheetName(drive.name)}`;
  const sheet = getOrCreateSheet(spreadsheet, sheetName);

  // ヘッダー設定
  setupDriveSheetHeaders(sheet);

  // フォルダ・ファイル構造取得
  const result = getDriveContents(drive.id);
  const items = result.items;
  const stats = result.stats;

  // データをシートに書き込み
  if (items.length > 0) {
    writeDriveDataToSheet(sheet, items);
  }

  // 統計情報を返す
  console.log(`処理完了: ${drive.name} (${items.length}件, ファイル: ${stats.totalFiles}, フォルダ: ${stats.totalFolders}, 容量: ${formatFileSize(stats.totalSize.toString())})`);

  return stats;
}

/**
 * ドライブの内容を再帰的に取得
 * @param {string} driveId ドライブID
 * @returns {Object} {items: Array, stats: Object} ファイル・フォルダ情報と統計
 */
function getDriveContents(driveId) {
  const items = [];
  const processedIds = new Set();
  const stats = {
    totalFiles: 0,
    totalFolders: 0,
    totalSize: 0,
    externalShareCount: 0
  };

  // 共有ドライブレベルの権限を取得
  let drivePermissions = [];
  try {
    const permissionsResponse = Drive.Permissions.list(driveId, {
      supportsAllDrives: true,
      useDomainAdminAccess: true,
      fields: 'permissions(id,type,role,emailAddress,displayName,domain,permissionDetails)'
    });
    drivePermissions = permissionsResponse.permissions || [];
    console.log(`共有ドライブレベル権限取得成功: ${drivePermissions.length}件`);
  } catch (error) {
    console.error(`共有ドライブレベル権限取得エラー (${driveId}):`, error);
    drivePermissions = [];
  }

  // ルートフォルダから開始
  collectItemsRecursive(driveId, driveId, '/', 0, items, processedIds, stats, drivePermissions);

  return { items, stats };
}

/**
 * アイテムを再帰的に収集
 * @param {string} driveId ドライブID
 * @param {string} parentId 親フォルダID
 * @param {string} currentPath 現在のパス
 * @param {number} level 階層レベル
 * @param {Array} items 収集結果配列
 * @param {Set} processedIds 処理済みID（無限ループ防止）
 * @param {Object} stats 統計情報オブジェクト
 * @param {Array} drivePermissions 共有ドライブレベルの権限配列
 */
function collectItemsRecursive(driveId, parentId, currentPath, level, items, processedIds, stats, drivePermissions) {
  // 階層制限チェック
  if (level > CONFIG.MAX_DEPTH) {
    console.warn(`最大階層(${CONFIG.MAX_DEPTH})に達しました: ${currentPath}`);
    return;
  }

  // 無限ループ防止
  if (processedIds.has(parentId)) {
    return;
  }
  processedIds.add(parentId);

  try {
    let pageToken = null;

    do {
      const params = {
        q: `'${parentId}' in parents and trashed = false`,
        pageSize: 1000,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        corpora: 'drive',
        driveId: driveId,
        useDomainAdminAccess: true,
        fields: 'nextPageToken,files(id,name,mimeType,parents,createdTime,modifiedTime,size,owners,lastModifyingUser,sharingUser,webViewLink)'
      };

      if (pageToken) {
        params.pageToken = pageToken;
      }

      const response = Drive.Files.list(params);

      if (response.files) {
        for (const file of response.files) {
          const itemPath = currentPath === '/' ? `/${file.name}` : `${currentPath}${file.name}`;
          const isFolder = file.mimeType === 'application/vnd.google-apps.folder';

          // 詳細な権限情報を取得（共有ドライブ対応）
          let detailedPermissions = [];
          let permissionError = null;

          try {
            // 共有ドライブのファイル権限を取得
            // 注意: useDomainAdminAccessはファイルレベルでは使用しない
            const permissionsResponse = Drive.Permissions.list(file.id, {
              supportsAllDrives: true,
              fields: 'permissions(id,type,role,emailAddress,displayName,domain,permissionDetails)'
            });
            detailedPermissions = permissionsResponse.permissions || [];

            // デバッグ用：権限数をログ出力
            if (detailedPermissions.length === 0) {
              console.log(`権限が0件: ${file.name} (${file.id})`);
            } else {
              console.log(`権限取得成功: ${file.name} - ${detailedPermissions.length}件`);
            }
          } catch (permError) {
            // 権限取得エラーをログに記録
            console.error(`権限取得エラー: ${file.name} (${file.id})`, permError);
            permissionError = permError.message;
            detailedPermissions = [];
          }

          // 権限レベル別にメンバーを分類
          const membersByRole = getDriveMembersByRole(detailedPermissions);

          // 共有ドライブレベルの権限を分類
          const driveMembers = getDriveMembersByRole(drivePermissions);

          // フォルダとファイルで表示する権限を調整
          let organizers, fileOrganizers, writers, editors, commenters, readers;

          if (isFolder) {
            // フォルダの場合：編集者は「ー」、その他は上位権限+個別権限を組み合わせ
            organizers = combinePermissions(driveMembers.organizers, membersByRole.organizers);
            fileOrganizers = combinePermissions(driveMembers.fileOrganizers, membersByRole.fileOrganizers);
            writers = combinePermissions(driveMembers.writers, membersByRole.writers);
            editors = 'ー';  // フォルダに編集者の概念はない
            commenters = combinePermissions(driveMembers.commenters, membersByRole.commenters);
            readers = combinePermissions(driveMembers.readers, membersByRole.readers);
          } else {
            // ファイルの場合：上位権限の管理者/コンテンツ管理者/投稿者も表示
            organizers = driveMembers.organizers.join(', ');
            fileOrganizers = driveMembers.fileOrganizers.join(', ');
            writers = driveMembers.writers.join(', ');
            editors = membersByRole.editors.join(', ');
            commenters = combinePermissions(driveMembers.commenters, membersByRole.commenters);
            readers = combinePermissions(driveMembers.readers, membersByRole.readers);
          }

          // アイテム情報を収集
          const itemInfo = {
            level: level,
            path: isFolder ? `${itemPath}/` : itemPath,
            type: isFolder ? 'フォルダ' : 'ファイル',
            name: file.name,
            id: file.id,
            parentId: parentId === driveId ? '' : parentId,
            creator: getCreatorName(file),
            createdTime: file.createdTime,
            modifiedTime: file.modifiedTime,
            size: formatFileSize(file.size),
            organizers: organizers,
            fileOrganizers: fileOrganizers,
            writers: writers,
            editors: editors,
            commenters: commenters,
            readers: readers,
            externalSharing: permissionError ? `エラー: ${permissionError}` : checkExternalSharing(detailedPermissions),
            url: file.webViewLink
          };

          items.push(itemInfo);

          // 統計情報を更新
          if (isFolder) {
            stats.totalFolders++;
            // フォルダの場合は再帰的に処理
            collectItemsRecursive(driveId, file.id, `${itemPath}/`, level + 1, items, processedIds, stats, drivePermissions);
          } else {
            stats.totalFiles++;
            const fileSize = parseInt(file.size) || 0;
            stats.totalSize += fileSize;
          }

          // 外部共有のカウント
          if (itemInfo.externalSharing !== 'なし') {
            stats.externalShareCount++;
          }
        }
      }

      pageToken = response.nextPageToken;

      // API制限対策
      if (CONFIG.API_DELAY > 0) {
        Utilities.sleep(CONFIG.API_DELAY);
      }

    } while (pageToken);

  } catch (error) {
    console.error(`フォルダ ${currentPath} の処理でエラー:`, error);
    logError(`フォルダ処理 (${currentPath})`, error.toString());
  }
}

// =============================================================================
// ユーティリティ関数
// =============================================================================

/**
 * 実行権限をチェック
 * @returns {boolean} 実行権限の有無
 */
function checkExecutionPermission() {
  const userEmail = Session.getActiveUser().getEmail();
  return CONFIG.ALLOWED_USERS.includes(userEmail);
}

/**
 * スプレッドシートを取得または作成
 * @returns {Spreadsheet} スプレッドシート
 */
function getOrCreateSpreadsheet() {
  try {
    const files = DriveApp.getFilesByName(CONFIG.SPREADSHEET_NAME);
    if (files.hasNext()) {
      return SpreadsheetApp.open(files.next());
    } else {
      return SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
    }
  } catch (error) {
    console.error('スプレッドシート取得エラー:', error);
    throw new Error('スプレッドシートの作成に失敗しました');
  }
}

/**
 * シートを取得または作成
 * @param {Spreadsheet} spreadsheet スプレッドシート
 * @param {string} sheetName シート名
 * @returns {Sheet} シート
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  return sheet;
}

/**
 * シート名をサニタイズ（Googleスプレッドシート制限対応）
 * @param {string} name 元の名前
 * @returns {string} サニタイズされた名前
 */
function sanitizeSheetName(name) {
  return name.replace(/[\[\]\/\\:*?"<>|]/g, '_').substring(0, 100);
}

/**
 * 作成者名を取得（共有ドライブ対応）
 * @param {Object} file ファイルオブジェクト
 * @returns {string} 作成者名
 */
function getCreatorName(file) {
  // 共有ユーザー（アップロードした人）を優先
  if (file.sharingUser) {
    return file.sharingUser.displayName || file.sharingUser.emailAddress || '共有者不明';
  }

  // 最終更新者
  if (file.lastModifyingUser) {
    return file.lastModifyingUser.displayName || file.lastModifyingUser.emailAddress || '更新者不明';
  }

  // オーナー（マイドライブの場合）
  if (file.owners && file.owners.length > 0) {
    return file.owners[0].displayName || file.owners[0].emailAddress || 'オーナー不明';
  }

  return '不明';
}

/**
 * ファイルサイズをフォーマット
 * @param {string} bytes バイト数
 * @returns {string} フォーマット済みサイズ
 */
function formatFileSize(bytes) {
  if (!bytes) return '';

  const size = parseInt(bytes);
  if (size === 0) return '0 B';

  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(size) / Math.log(k));

  return parseFloat((size / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * 権限情報のサマリーを作成
 * @param {Array} permissions 権限配列
 * @returns {string} 権限サマリー
 */
function getPermissionsSummary(permissions) {
  if (!permissions || permissions.length === 0) return '権限なし';

  const details = [];
  for (const permission of permissions) {
    const role = permission.role || '不明';
    const type = permission.type || '不明';
    let identifier = '';

    if (permission.emailAddress) {
      identifier = permission.emailAddress;
    } else if (permission.displayName) {
      identifier = permission.displayName;
    } else if (permission.domain) {
      identifier = `@${permission.domain}`;
    } else if (type === 'anyone') {
      identifier = '全員';
    }

    details.push(`${identifier}(${role})`);
  }

  return details.join(', ');
}

/**
 * 権限を持つ全メンバーをリスト化（表示名:メールアドレス形式）
 * @param {Array} permissions 権限配列
 * @returns {string} メンバーリスト（カンマ区切り）
 */
function getAllMembers(permissions) {
  if (!permissions || permissions.length === 0) {
    return '';
  }

  const members = [];

  for (const permission of permissions) {
    const type = permission.type || '';
    let identifier = '';

    // メールアドレスと表示名がある場合
    if (permission.emailAddress && permission.displayName) {
      identifier = `${permission.displayName}:${permission.emailAddress}`;
    } else if (permission.emailAddress) {
      identifier = permission.emailAddress;
    } else if (permission.displayName) {
      identifier = permission.displayName;
    } else if (permission.domain) {
      identifier = `@${permission.domain}`;
    } else if (type === 'anyone') {
      identifier = '全員';
    }

    if (identifier) {
      members.push(identifier);
    }
  }

  return members.join(', ');
}

/**
 * 上位権限と個別権限を組み合わせて重複を除去
 * @param {Array} upperPermissions 上位権限（共有ドライブレベル）の配列
 * @param {Array} individualPermissions 個別権限の配列
 * @returns {string} 組み合わせた権限文字列（カンマ区切り）
 */
function combinePermissions(upperPermissions, individualPermissions) {
  const combined = new Set();

  // 上位権限を追加（先に追加して優先表示）
  if (upperPermissions && upperPermissions.length > 0) {
    upperPermissions.forEach(permission => {
      if (permission && permission.trim()) {
        combined.add(permission.trim());
      }
    });
  }

  // 個別権限を追加
  if (individualPermissions && individualPermissions.length > 0) {
    individualPermissions.forEach(permission => {
      if (permission && permission.trim()) {
        combined.add(permission.trim());
      }
    });
  }

  return Array.from(combined).join(', ');
}

/**
 * 共有ドライブのメンバーを権限レベル別に分類
 * @param {Array} permissions 権限配列
 * @returns {Object} 権限レベル別のメンバーリスト
 */
function getDriveMembersByRole(permissions) {
  const result = {
    organizers: [],        // 管理者 (organizer)
    fileOrganizers: [],    // コンテンツ管理者 (fileOrganizer)
    writers: [],           // 投稿者 (writer)
    editors: [],           // 編集者 (ファイルのwriter相当)
    commenters: [],        // 閲覧者（コメント可） (commenter)
    readers: []            // 閲覧者 (reader)
  };

  if (!permissions || permissions.length === 0) {
    return result;
  }

  for (const permission of permissions) {
    const role = permission.role || '';
    const type = permission.type || '';
    let identifier = '';

    // メールアドレスと表示名がある場合
    if (permission.emailAddress && permission.displayName) {
      identifier = `${permission.displayName}:${permission.emailAddress}`;
    } else if (permission.emailAddress) {
      identifier = permission.emailAddress;
    } else if (permission.displayName) {
      identifier = permission.displayName;
    } else if (permission.domain) {
      identifier = `@${permission.domain}`;
    } else if (type === 'anyone') {
      identifier = '全員';
    }

    if (!identifier) continue;

    // 権限レベルごとに分類
    if (role === 'organizer') {
      result.organizers.push(identifier);
    } else if (role === 'fileOrganizer') {
      result.fileOrganizers.push(identifier);
    } else if (role === 'writer') {
      // writerは投稿者と編集者の両方に使われる可能性があるため、両方に格納
      result.writers.push(identifier);
      result.editors.push(identifier);
    } else if (role === 'commenter') {
      result.commenters.push(identifier);
    } else if (role === 'reader') {
      result.readers.push(identifier);
    }
  }

  return result;
}

/**
 * メールアドレスが内部ドメインかどうかチェック
 * @param {string} email メールアドレス
 * @returns {boolean} 内部ドメインの場合true
 */
function isInternalDomain(email) {
  if (!email) return false;
  return CONFIG.COMPANY_DOMAINS.some(domain => email.endsWith(`@${domain}`));
}

/**
 * 外部共有をチェック（組織内ドメイン共有と区別）
 * @param {Array} permissions 権限配列
 * @returns {string} 外部共有状況
 */
function checkExternalSharing(permissions) {
  if (!permissions || permissions.length === 0) return 'なし';

  let internalDomainShare = false;
  const externalUsers = [];

  for (const p of permissions) {
    // 組織内ドメイン全体への共有（type=domain でドメインが会社ドメインの場合）
    if (p.type === 'domain' && CONFIG.COMPANY_DOMAINS.includes(p.domain)) {
      internalDomainShare = true;
      continue;
    }

    // 真の外部共有をチェック
    if (p.type === 'anyone') {
      externalUsers.push(p);
    } else if (p.type === 'domain' && !CONFIG.COMPANY_DOMAINS.includes(p.domain)) {
      externalUsers.push(p);
    } else if (p.emailAddress && !isInternalDomain(p.emailAddress)) {
      externalUsers.push(p);
    }
  }

  // 結果を組み立てる
  const parts = [];
  if (internalDomainShare) {
    parts.push('組織内共有あり');
  }
  if (externalUsers.length > 0) {
    parts.push(`外部共有あり(${externalUsers.length}件)`);
  }

  return parts.length > 0 ? parts.join(', ') : 'なし';
}

/**
 * 実行時間制限チェック
 * @param {Date} startTime 開始時刻
 * @returns {boolean} タイムアウト接近中かどうか
 */
function isTimeoutApproaching(startTime) {
  const elapsedTime = (new Date() - startTime) / 1000;
  return elapsedTime > CONFIG.MAX_EXECUTION_TIME;
}

/**
 * 実行状況を更新
 * @param {string} status ステータス
 * @param {number} progress 進捗率（0-100）
 */
function updateExecutionStatus(status, progress) {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('CURRENT_SPREADSHEET_ID');
    if (!spreadsheetId) {
      console.warn('スプレッドシートIDが未設定のため、ステータス更新をスキップします');
      return;
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const controlSheet = spreadsheet.getSheetByName('00_実行管理');
    if (controlSheet) {
      controlSheet.getRange('B2').setValue(status);
      controlSheet.getRange('B3').setValue(`${Math.round(progress)}%`);
      controlSheet.getRange('B4').setValue(new Date());
    }
  } catch (error) {
    console.warn('ステータス更新エラー:', error);
  }
}

/**
 * エラーをログに記録
 * @param {string} context コンテキスト
 * @param {string} errorMessage エラーメッセージ
 */
function logError(context, errorMessage) {
  console.error(`[${context}] ${errorMessage}`);

  try {
    const spreadsheet = getOrCreateSpreadsheet();
    const logSheet = getOrCreateSheet(spreadsheet, '99_エラーログ');

    // ヘッダーが未設定の場合
    if (logSheet.getLastRow() === 0) {
      logSheet.getRange(1, 1, 1, 3).setValues([['日時', 'コンテキスト', 'エラー内容']]);
    }

    const newRow = logSheet.getLastRow() + 1;
    logSheet.getRange(newRow, 1, 1, 3).setValues([[new Date(), context, errorMessage]]);

  } catch (logError) {
    console.error('ログ記録エラー:', logError);
  }
}

// =============================================================================
// スプレッドシート作成・更新
// =============================================================================

/**
 * マスターシートを作成・更新
 * @param {Spreadsheet} spreadsheet スプレッドシート
 * @param {Array} sharedDrives 共有ドライブリスト
 */
function createMasterSheet(spreadsheet, sharedDrives) {
  const sheet = getOrCreateSheet(spreadsheet, '00_共有ドライブ一覧');

  // ヘッダー設定
  const headers = [
    'No', 'ドライブ名', 'ドライブID', '作成日', 'ファイル数', '容量(GB)',
    '最終更新', '対応シート', '管理者', 'コンテンツ管理者', '投稿者', '閲覧者（コメント可）', '閲覧者', '外部共有', '状況', 'URL'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 共有ドライブのメンバー情報と外部共有状況を取得
  const driveMembers = {};
  const driveExternalSharing = {};
  for (const drive of sharedDrives) {
    try {
      const permissions = Drive.Permissions.list(drive.id, {
        supportsAllDrives: true,
        useDomainAdminAccess: true,
        fields: 'permissions(id,type,role,emailAddress,displayName,domain)'
      });
      const perms = permissions.permissions || [];
      driveMembers[drive.id] = getDriveMembersByRole(perms);
      driveExternalSharing[drive.id] = checkExternalSharing(perms);
    } catch (error) {
      console.warn(`メンバー情報取得エラー (${drive.name}):`, error.message);
      driveMembers[drive.id] = {
        organizers: [],
        fileOrganizers: [],
        writers: [],
        commenters: [],
        readers: []
      };
      driveExternalSharing[drive.id] = 'なし';
    }
    Utilities.sleep(CONFIG.API_DELAY);
  }

  // データ設定
  const data = sharedDrives.map((drive, index) => {
    const members = driveMembers[drive.id] || {
      organizers: [],
      fileOrganizers: [],
      writers: [],
      commenters: [],
      readers: []
    };

    return [
      index + 1,
      drive.name,
      drive.id,
      drive.createdTime ? new Date(drive.createdTime) : '',
      '', // ファイル数（後で更新）
      '', // 容量（後で更新）
      '', // 最終更新（後で更新）
      `${String(index + 1).padStart(2, '0')}_${sanitizeSheetName(drive.name)}`,
      members.organizers.join(', '),        // 管理者
      members.fileOrganizers.join(', '),    // コンテンツ管理者
      members.writers.join(', '),           // 投稿者
      members.commenters.join(', '),        // 閲覧者（コメント可）
      members.readers.join(', '),           // 閲覧者
      driveExternalSharing[drive.id] || 'なし', // 外部共有
      '未処理',
      `https://drive.google.com/drive/folders/${drive.id}`
    ];
  });

  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }

  // 書式設定
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * ドライブシートのヘッダーを設定
 * @param {Sheet} sheet シート
 */
function setupDriveSheetHeaders(sheet) {
  const headers = [
    'レベル', 'パス', '種別', '名前', 'ID', '親ID', '作成者', '作成日',
    '更新日', 'サイズ', '管理者', 'コンテンツ管理者', '投稿者', '編集者', '閲覧者（コメント可）', '閲覧者', '外部共有', 'URL'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#34a853').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/**
 * ドライブデータをシートに書き込み
 * @param {Sheet} sheet シート
 * @param {Array} items アイテムリスト
 */
function writeDriveDataToSheet(sheet, items) {
  const data = items.map(item => [
    item.level,
    item.path,
    item.type,
    item.name,
    item.id,
    item.parentId,
    item.creator,
    item.createdTime ? new Date(item.createdTime) : '',
    item.modifiedTime ? new Date(item.modifiedTime) : '',
    item.size,
    item.organizers || '',
    item.fileOrganizers || '',
    item.writers || '',
    item.editors || '',
    item.commenters || '',
    item.readers || '',
    item.externalSharing,
    item.url
  ]);

  // バッチ書き込み
  for (let i = 0; i < data.length; i += CONFIG.BATCH_SIZE) {
    const batch = data.slice(i, i + CONFIG.BATCH_SIZE);
    const startRow = i + 2; // ヘッダー行の次から
    sheet.getRange(startRow, 1, batch.length, batch[0].length).setValues(batch);
  }

  // 列幅自動調整
  sheet.autoResizeColumns(1, 18);
}

/**
 * マスターシートを統計情報で更新
 * @param {Spreadsheet} spreadsheet スプレッドシート
 * @param {Array} sharedDrives 共有ドライブリスト
 * @param {Object} driveStats ドライブ統計情報
 */
function updateMasterSheetWithStats(spreadsheet, sharedDrives, driveStats) {
  const sheet = spreadsheet.getSheetByName('00_共有ドライブ一覧');
  if (!sheet) return;

  for (let i = 0; i < sharedDrives.length; i++) {
    const drive = sharedDrives[i];
    const stats = driveStats[drive.id];

    if (stats) {
      const row = i + 2; // ヘッダー行の次から

      // ファイル数 (E列)
      sheet.getRange(row, 5).setValue(stats.totalFiles);

      // 容量(GB) (F列)
      const sizeInGB = (stats.totalSize / (1024 * 1024 * 1024)).toFixed(2);
      sheet.getRange(row, 6).setValue(parseFloat(sizeInGB));

      // 外部共有 (N列) - 既存の値（共有ドライブメンバーレベル）を取得
      const currentExternalStatus = sheet.getRange(row, 14).getValue() || '';
      let externalStatus = '';

      // ファイルレベルの外部共有をチェック
      const fileExternalCount = stats.externalShareCount || 0;

      // 既に共有ドライブメンバーレベルで外部共有が検出されている場合
      if (currentExternalStatus && currentExternalStatus !== 'なし') {
        if (fileExternalCount > 0) {
          externalStatus = `${currentExternalStatus}, ファイルレベル外部共有あり(${fileExternalCount}件)`;
        } else {
          externalStatus = currentExternalStatus;
        }
      } else {
        // 共有ドライブメンバーレベルで外部共有がない場合
        externalStatus = fileExternalCount > 0 ? `ファイルレベル外部共有あり(${fileExternalCount}件)` : 'なし';
      }

      sheet.getRange(row, 14).setValue(externalStatus);

      // 状況 (O列)
      sheet.getRange(row, 15).setValue('完了');
    }
  }
}

/**
 * 共有ドライブレベルの権限情報を取得
 * @param {string} driveId ドライブID
 * @returns {string} 権限サマリー
 */
function getDrivePermissions(driveId) {
  try {
    const permissions = Drive.Permissions.list(driveId, {
      supportsAllDrives: true,
      fields: 'permissions(id,type,role,emailAddress,displayName,domain)'
    });

    if (!permissions.permissions || permissions.permissions.length === 0) {
      return '権限なし';
    }

    const details = [];
    for (const permission of permissions.permissions) {
      const role = permission.role || '不明';
      let identifier = '';

      if (permission.emailAddress) {
        identifier = permission.emailAddress;
      } else if (permission.displayName) {
        identifier = permission.displayName;
      } else if (permission.domain) {
        identifier = `@${permission.domain}`;
      } else if (permission.type === 'anyone') {
        identifier = '全員';
      }

      details.push(`${identifier}(${role})`);
    }

    return details.join(', ');
  } catch (error) {
    console.warn(`ドライブ権限取得エラー (${driveId}):`, error);
    return '取得失敗';
  }
}

/**
 * 進捗を保存
 * @param {number} currentIndex 現在のインデックス
 * @param {Array} sharedDrives 共有ドライブリスト
 */
function saveProgress(currentIndex, sharedDrives) {
  PropertiesService.getScriptProperties().setProperties({
    'CURRENT_INDEX': currentIndex.toString(),
    'TOTAL_DRIVES': sharedDrives.length.toString(),
    'SHARED_DRIVES': JSON.stringify(sharedDrives)
  });
}

/**
 * 保存された進捗を取得
 * @returns {Object|null} 進捗情報
 */
function getProgressFromProperties() {
  const properties = PropertiesService.getScriptProperties();
  const currentIndex = properties.get('CURRENT_INDEX');

  if (currentIndex === null) return null;

  return {
    currentIndex: parseInt(currentIndex),
    totalDrives: parseInt(properties.get('TOTAL_DRIVES')),
    sharedDrives: JSON.parse(properties.get('SHARED_DRIVES'))
  };

}
