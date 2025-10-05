/**
 * Google Drive 共有ドライブ情報収集ツール v2.2
 * 全共有ドライブのフォルダ・ファイル構造と権限情報、共有設定を取得
 */

const CONFIG = {
  SPREADSHEET_NAME: 'Drive情報収集結果',
  BATCH_SIZE: 1000,
  API_DELAY: 100,
  MAX_EXECUTION_TIME: 330,
  MAX_FILES_PER_BATCH: 1000,
  MAX_DEPTH: 10,
  ALLOWED_USERS: ['*****@*****.co.jp'],
  COMPANY_DOMAINS: ['******.co.jp']
};

// スプレッドシート列番号定数
const SHEET_COLUMNS = {
  MASTER: {
    FILE_COUNT: 5,
    SIZE_GB: 6,
    SHEET_NAME: 8,
    EXTERNAL_SHARE: 14,
    DOMAIN_USERS_ONLY: 15,      // 組織外アクセス
    DRIVE_MEMBERS_ONLY: 16,     // メンバー外アクセス
    SHARING_FOLDERS: 17,        // フォルダ共有(コンテンツ管理者)
    COPY_RESTRICTION: 18,       // コピー制限
    STATUS: 19                  // 状況
  },
  DATA_ROW_START: 2
};

// =============================================================================
// メイン実行関数
// =============================================================================

function driveGet() {
  try {
    console.log('=== driveGet: 共有ドライブ一覧取得開始 ===');

    if (!checkExecutionPermission()) {
      throw new Error('実行権限がありません。管理者にお問い合わせください。');
    }

    const spreadsheet = getOrCreateSpreadsheet();
    PropertiesService.getScriptProperties().setProperty('CURRENT_SPREADSHEET_ID', spreadsheet.getId());

    const sharedDrives = getSharedDrives();
    console.log(`共有ドライブ数: ${sharedDrives.length}`);

    createMasterSheet(spreadsheet, sharedDrives);

    console.log('=== driveGet: 完了 ===');
    console.log('次に fileGet() を実行してファイル情報を取得してください。');

  } catch (error) {
    console.error('driveGet処理でエラー:', error);
    throw error;
  }
}

function fileGet() {
  try {
    console.log('=== fileGet: ファイル情報取得開始 ===');

    if (!checkExecutionPermission()) {
      throw new Error('実行権限がありません。管理者にお問い合わせください。');
    }

    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('CURRENT_SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが見つかりません。先に driveGet() を実行してください。');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const masterSheet = spreadsheet.getSheetByName('00_共有ドライブ一覧');

    if (!masterSheet) {
      throw new Error('00_共有ドライブ一覧シートが見つかりません。先に driveGet() を実行してください。');
    }

    const progress = getCurrentProgress();
    let targetRow, targetDrive, isContinuation, sheet;

    if (progress && progress.driveId) {
      console.log(`継続処理: ${progress.driveName} (処理済み: ${progress.processedCount}件)`);

      targetDrive = {
        id: progress.driveId,
        name: progress.driveName,
        sheetName: progress.sheetName
      };
      targetRow = progress.masterSheetRow;
      isContinuation = true;

      sheet = spreadsheet.getSheetByName(targetDrive.sheetName);
      if (!sheet) {
        throw new Error(`シート ${targetDrive.sheetName} が見つかりません。`);
      }

    } else {
      const lastRow = masterSheet.getLastRow();
      if (lastRow < SHEET_COLUMNS.DATA_ROW_START) {
        throw new Error('共有ドライブ一覧にデータがありません。');
      }

      const statusColumn = SHEET_COLUMNS.MASTER.STATUS;
      const data = masterSheet.getRange(SHEET_COLUMNS.DATA_ROW_START, 1, lastRow - 1, statusColumn).getValues();

      for (let i = 0; i < data.length; i++) {
        const status = data[i][statusColumn - 1];
        if (status === '未処理' || status === '') {
          targetRow = i + SHEET_COLUMNS.DATA_ROW_START;
          targetDrive = {
            index: data[i][0],
            name: data[i][1],
            id: data[i][2],
            sheetName: data[i][SHEET_COLUMNS.MASTER.SHEET_NAME - 1]
          };
          break;
        }
      }

      if (!targetDrive) {
        console.log('全ての共有ドライブの処理が完了しています。');
        Browser.msgBox('完了', '全ての共有ドライブの処理が完了しています。', Browser.Buttons.OK);
        return;
      }

      console.log(`新規処理開始: ${targetDrive.name} (ID: ${targetDrive.id})`);

      sheet = getOrCreateSheet(spreadsheet, targetDrive.sheetName);
      setupDriveSheetHeaders(sheet);
    }

    const result = getDriveContentsPaginated(targetDrive.id, progress);
    const items = result.items;
    const stats = result.stats;
    const hasMore = result.hasMore;
    const nextProgress = result.nextProgress;

    if (items.length > 0) {
      const startRow = isContinuation ? progress.sheetLastRow : SHEET_COLUMNS.DATA_ROW_START;
      writeDriveDataToSheetAppend(sheet, items, startRow);
    }

    if (hasMore) {
      const progressData = {
        driveId: targetDrive.id,
        driveName: targetDrive.name,
        sheetName: targetDrive.sheetName,
        masterSheetRow: targetRow,
        processedCount: nextProgress.processedCount,
        sheetLastRow: nextProgress.sheetLastRow,
        stats: nextProgress.stats,
        recursionState: nextProgress.recursionState
      };

      saveProgress(progressData);
      masterSheet.getRange(targetRow, SHEET_COLUMNS.MASTER.STATUS).setValue(`処理中 (${nextProgress.processedCount}件)`);

      console.log(`=== fileGet: ${targetDrive.name} 一時停止 (${nextProgress.processedCount}件処理) ===`);
      console.log(`再度 fileGet() を実行して続きを処理してください。`);

    } else {
      clearProgress();
      updateMasterSheetRow(masterSheet, targetRow, stats);

      console.log(`=== fileGet: ${targetDrive.name} 処理完了 ===`);
      console.log(`ファイル: ${stats.totalFiles}, フォルダ: ${stats.totalFolders}, 容量: ${formatFileSize(stats.totalSize.toString())}`);

      const lastRow = masterSheet.getLastRow();
      const statusColumn = SHEET_COLUMNS.MASTER.STATUS;
      const data = masterSheet.getRange(SHEET_COLUMNS.DATA_ROW_START, 1, lastRow - 1, statusColumn).getValues();
      const remaining = data.filter(row => row[statusColumn - 1] === '未処理' || row[statusColumn - 1] === '').length;

      if (remaining > 0) {
        console.log(`残り ${remaining} 件の共有ドライブが未処理です。再度 fileGet() を実行してください。`);
      } else {
        console.log('全ての共有ドライブの処理が完了しました!');
      }
    }

  } catch (error) {
    console.error('fileGet処理でエラー:', error);
    const progress = getCurrentProgress();
    if (progress) {
      console.log('進捗は保持されています。エラー修正後、再度 fileGet() を実行してください。');
    }
    throw error;
  }
}

// =============================================================================
// 進捗管理
// =============================================================================

function getCurrentProgress() {
  const props = PropertiesService.getScriptProperties();
  const progressJson = props.getProperty('CURRENT_PROGRESS');
  if (!progressJson) return null;

  try {
    return JSON.parse(progressJson);
  } catch (error) {
    return null;
  }
}

function saveProgress(progress) {
  PropertiesService.getScriptProperties().setProperty('CURRENT_PROGRESS', JSON.stringify(progress));
}

function clearProgress() {
  PropertiesService.getScriptProperties().deleteProperty('CURRENT_PROGRESS');
}

// =============================================================================
// ドライブ情報取得
// =============================================================================

function getSharedDrives() {
  try {
    const drives = [];
    let pageToken = null;

    do {
      const params = {
        pageSize: 100,
        fields: 'nextPageToken,drives(id,name,createdTime,capabilities,restrictions)',
        useDomainAdminAccess: true
      };

      if (pageToken) params.pageToken = pageToken;

      const response = Drive.Drives.list(params);
      if (response.drives) drives.push(...response.drives);
      pageToken = response.nextPageToken;

    } while (pageToken);

    return drives;

  } catch (error) {
    throw new Error(`共有ドライブの取得に失敗しました: ${error.message}`);
  }
}

function getDriveContentsPaginated(driveId, progress) {
  const items = [];
  let processedIds = new Set();
  let stats = {
    totalFiles: 0,
    totalFolders: 0,
    totalSize: 0,
    externalShareCount: 0
  };

  let recursionState = null;

  if (progress && progress.recursionState) {
    recursionState = progress.recursionState;
    processedIds = new Set(recursionState.processedIds || []);
    stats = progress.stats || stats;
  }

  let drivePermissions = [];
  try {
    const permissionsResponse = Drive.Permissions.list(driveId, {
      supportsAllDrives: true,
      useDomainAdminAccess: true,
      fields: 'permissions(id,type,role,emailAddress,displayName,domain,permissionDetails)'
    });
    drivePermissions = permissionsResponse.permissions || [];
  } catch (error) {
    console.error(`共有ドライブレベル権限取得エラー (${driveId}):`, error);
  }

  // 再開時の親ID、パス、レベルを決定
  let startParentId = driveId;
  let startPath = '/';
  let startLevel = 0;

  if (recursionState) {
    if (recursionState.parentId === null) {
      startParentId = null;
    } else if (recursionState.parentId) {
      startParentId = recursionState.parentId;
    }

    if (recursionState.currentPath) {
      startPath = recursionState.currentPath;
    }

    if (recursionState.level !== null && recursionState.level !== undefined) {
      startLevel = recursionState.level;
    }
  }

  const result = collectItemsRecursivePaginated(
    driveId,
    startParentId,
    startPath,
    startLevel,
    items,
    processedIds,
    stats,
    drivePermissions,
    CONFIG.MAX_FILES_PER_BATCH,
    recursionState
  );

  const nextProgress = {
    processedCount: (progress ? progress.processedCount : 0) + items.length,
    sheetLastRow: (progress ? progress.sheetLastRow : SHEET_COLUMNS.DATA_ROW_START) + items.length,
    stats: stats,
    recursionState: result.hasMore ? result.recursionState : null
  };

  return {
    items,
    stats,
    hasMore: result.hasMore,
    nextProgress
  };
}

function collectItemsRecursivePaginated(driveId, parentId, currentPath, level, items, processedIds, stats, drivePermissions, maxItems, recursionState) {

  if (parentId === null) {
    let pageToken = null;
    let folderQueue = [];

    if (recursionState) {
      pageToken = recursionState.pageToken || null;
      folderQueue = recursionState.folderQueue || [];

      if (folderQueue.length > 0) {
        let firstFolderState = recursionState.nestedState || null;

        while (folderQueue.length > 0) {
          if (items.length >= maxItems) {
            return {
              hasMore: true,
              recursionState: {
                parentId: null,
                currentPath: null,
                level: null,
                pageToken: null,
                folderQueue: folderQueue,
                processedIds: Array.from(processedIds)
              }
            };
          }

          const nextFolder = folderQueue.shift();
          const result = collectItemsRecursivePaginated(
            driveId,
            nextFolder.id,
            nextFolder.path,
            nextFolder.level,
            items,
            processedIds,
            stats,
            drivePermissions,
            maxItems,
            firstFolderState
          );

          firstFolderState = null;

          if (result.hasMore) {
            return {
              hasMore: true,
              recursionState: {
                parentId: null,
                currentPath: null,
                level: null,
                pageToken: null,
                folderQueue: folderQueue,
                processedIds: Array.from(processedIds),
                nestedState: result.recursionState
              }
            };
          }
        }

        return { hasMore: false, recursionState: null };
      }
    }

    throw new Error('parentIdがnullですが、復元可能な状態が見つかりません');
  }

  if (level > CONFIG.MAX_DEPTH) {
    return { hasMore: false, recursionState: null };
  }

  if (processedIds.has(parentId)) {
    return { hasMore: false, recursionState: null };
  }
  processedIds.add(parentId);

  let pageToken = null;
  let folderQueue = [];

  if (recursionState) {
    pageToken = recursionState.pageToken || null;
    folderQueue = recursionState.folderQueue || [];
  }

  try {
    do {
      if (items.length >= maxItems) {
        return {
          hasMore: true,
          recursionState: {
            parentId: parentId,
            currentPath: currentPath,
            level: level,
            pageToken: pageToken,
            folderQueue: folderQueue,
            processedIds: Array.from(processedIds)
          }
        };
      }

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

      if (pageToken) params.pageToken = pageToken;

      const response = Drive.Files.list(params);

      if (response.files) {
        for (const file of response.files) {

          if (items.length >= maxItems) {
            return {
              hasMore: true,
              recursionState: {
                parentId: parentId,
                currentPath: currentPath,
                level: level,
                pageToken: response.nextPageToken || null,
                folderQueue: folderQueue,
                processedIds: Array.from(processedIds)
              }
            };
          }

          const itemPath = currentPath === '/' ? `/${file.name}` : `${currentPath}${file.name}`;
          const isFolder = file.mimeType === 'application/vnd.google-apps.folder';

          let detailedPermissions = [];
          let permissionError = null;

          try {
            const permissionsResponse = Drive.Permissions.list(file.id, {
              supportsAllDrives: true,
              fields: 'permissions(id,type,role,emailAddress,displayName,domain,permissionDetails)'
            });
            detailedPermissions = permissionsResponse.permissions || [];
          } catch (permError) {
            permissionError = permError.message;
          }

          const membersByRole = getDriveMembersByRole(detailedPermissions);
          const driveMembers = getDriveMembersByRole(drivePermissions);

          let organizers, fileOrganizers, writers, editors, commenters, readers;

          if (isFolder) {
            organizers = combinePermissions(driveMembers.organizers, membersByRole.organizers);
            fileOrganizers = combinePermissions(driveMembers.fileOrganizers, membersByRole.fileOrganizers);
            writers = combinePermissions(driveMembers.writers, membersByRole.writers);
            editors = 'ー';
            commenters = combinePermissions(driveMembers.commenters, membersByRole.commenters);
            readers = combinePermissions(driveMembers.readers, membersByRole.readers);
          } else {
            organizers = driveMembers.organizers.join(', ');
            fileOrganizers = driveMembers.fileOrganizers.join(', ');
            writers = driveMembers.writers.join(', ');
            editors = membersByRole.editors.join(', ');
            commenters = combinePermissions(driveMembers.commenters, membersByRole.commenters);
            readers = combinePermissions(driveMembers.readers, membersByRole.readers);
          }

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

          if (isFolder) {
            stats.totalFolders++;
            folderQueue.push({
              id: file.id,
              path: `${itemPath}/`,
              level: level + 1
            });
          } else {
            stats.totalFiles++;
            const fileSize = parseInt(file.size) || 0;
            stats.totalSize += fileSize;
          }

          if (itemInfo.externalSharing !== 'なし') {
            stats.externalShareCount++;
          }
        }
      }

      pageToken = response.nextPageToken;

      if (CONFIG.API_DELAY > 0) {
        Utilities.sleep(CONFIG.API_DELAY);
      }

    } while (pageToken);

    while (folderQueue.length > 0) {

      if (items.length >= maxItems) {
        return {
          hasMore: true,
          recursionState: {
            parentId: null,
            currentPath: null,
            level: null,
            pageToken: null,
            folderQueue: folderQueue,
            processedIds: Array.from(processedIds)
          }
        };
      }

      const nextFolder = folderQueue.shift();
      const result = collectItemsRecursivePaginated(
        driveId,
        nextFolder.id,
        nextFolder.path,
        nextFolder.level,
        items,
        processedIds,
        stats,
        drivePermissions,
        maxItems,
        null
      );

      if (result.hasMore) {
        return {
          hasMore: true,
          recursionState: {
            parentId: null,
            currentPath: null,
            level: null,
            pageToken: null,
            folderQueue: folderQueue,
            processedIds: Array.from(processedIds),
            nestedState: result.recursionState
          }
        };
      }
    }

    return { hasMore: false, recursionState: null };

  } catch (error) {
    console.error(`フォルダ ${currentPath} の処理でエラー:`, error);
    return { hasMore: false, recursionState: null };
  }
}

// =============================================================================
// ユーティリティ
// =============================================================================

function checkExecutionPermission() {
  const userEmail = Session.getActiveUser().getEmail();
  return CONFIG.ALLOWED_USERS.includes(userEmail);
}

function getOrCreateSpreadsheet() {
  try {
    const files = DriveApp.getFilesByName(CONFIG.SPREADSHEET_NAME);
    if (files.hasNext()) {
      return SpreadsheetApp.open(files.next());
    } else {
      return SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
    }
  } catch (error) {
    throw new Error('スプレッドシートの作成に失敗しました');
  }
}

function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  return sheet;
}

function sanitizeSheetName(name) {
  return name.replace(/[\[\]\/\\:*?"<>|]/g, '_').substring(0, 100);
}

function getCreatorName(file) {
  if (file.sharingUser) {
    return file.sharingUser.displayName || file.sharingUser.emailAddress || '共有者不明';
  }

  if (file.lastModifyingUser) {
    return file.lastModifyingUser.displayName || file.lastModifyingUser.emailAddress || '更新者不明';
  }

  if (file.owners && file.owners.length > 0) {
    return file.owners[0].displayName || file.owners[0].emailAddress || 'オーナー不明';
  }

  return '不明';
}

function formatFileSize(bytes) {
  if (!bytes) return '';

  const size = parseInt(bytes);
  if (size === 0) return '0 B';

  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(size) / Math.log(k));

  return parseFloat((size / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function combinePermissions(upperPermissions, individualPermissions) {
  const combined = new Set();

  if (upperPermissions && upperPermissions.length > 0) {
    upperPermissions.forEach(permission => {
      if (permission && permission.trim()) {
        combined.add(permission.trim());
      }
    });
  }

  if (individualPermissions && individualPermissions.length > 0) {
    individualPermissions.forEach(permission => {
      if (permission && permission.trim()) {
        combined.add(permission.trim());
      }
    });
  }

  return Array.from(combined).join(', ');
}

function getDriveMembersByRole(permissions) {
  const result = {
    organizers: [],
    fileOrganizers: [],
    writers: [],
    editors: [],
    commenters: [],
    readers: []
  };

  if (!permissions || permissions.length === 0) return result;

  for (const permission of permissions) {
    const role = permission.role || '';
    const type = permission.type || '';
    let identifier = '';

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

    if (role === 'organizer') {
      result.organizers.push(identifier);
    } else if (role === 'fileOrganizer') {
      result.fileOrganizers.push(identifier);
    } else if (role === 'writer') {
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

function isInternalDomain(email) {
  if (!email) return false;
  return CONFIG.COMPANY_DOMAINS.some(domain => email.endsWith(`@${domain}`));
}

function checkExternalSharing(permissions) {
  if (!permissions || permissions.length === 0) return 'なし';

  let internalDomainShare = false;
  const externalUsers = [];

  for (const p of permissions) {
    if (p.type === 'domain' && CONFIG.COMPANY_DOMAINS.includes(p.domain)) {
      internalDomainShare = true;
      continue;
    }

    if (p.type === 'anyone') {
      externalUsers.push(p);
    } else if (p.type === 'domain' && !CONFIG.COMPANY_DOMAINS.includes(p.domain)) {
      externalUsers.push(p);
    } else if (p.emailAddress && !isInternalDomain(p.emailAddress)) {
      externalUsers.push(p);
    }
  }

  const parts = [];
  if (internalDomainShare) {
    parts.push('組織内共有あり');
  }
  if (externalUsers.length > 0) {
    parts.push(`外部共有あり(${externalUsers.length}件)`);
  }

  return parts.length > 0 ? parts.join(', ') : 'なし';
}

// 空のメンバー情報オブジェクトを生成
function createEmptyMembers() {
  return {
    organizers: [],
    fileOrganizers: [],
    writers: [],
    commenters: [],
    readers: []
  };
}

// =============================================================================
// スプレッドシート操作
// =============================================================================

function createMasterSheet(spreadsheet, sharedDrives) {
  const sheet = getOrCreateSheet(spreadsheet, '00_共有ドライブ一覧');

  const headers = [
    'No', 'ドライブ名', 'ドライブID', '作成日', 'ファイル数', '容量(GB)',
    '最終更新', '対応シート', '管理者', 'コンテンツ管理者', '投稿者', '閲覧者(コメント可)', '閲覧者', '外部共有',
    '組織外アクセス', 'メンバー外アクセス', 'フォルダ共有(コンテンツ管理者)', 'コピー制限',
    '状況', 'URL'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

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
      driveMembers[drive.id] = createEmptyMembers();
      driveExternalSharing[drive.id] = 'なし';
    }
    Utilities.sleep(CONFIG.API_DELAY);
  }

  const data = sharedDrives.map((drive, index) => {
    const members = driveMembers[drive.id] || createEmptyMembers();
    const restrictions = drive.restrictions || {};

    // 設定情報の判定
    const domainUsersOnly = restrictions.domainUsersOnly ? '禁止' : '許可';
    const driveMembersOnly = restrictions.driveMembersOnly ? '禁止' : '許可';
    const sharingFoldersRequiresOrganizer = restrictions.sharingFoldersRequiresOrganizerPermission ? '管理者のみ' : '許可';
    const copyRequiresWriter = restrictions.copyRequiresWriterPermission ? '投稿者以上' : '全員可';

    return [
      index + 1,
      drive.name,
      drive.id,
      drive.createdTime ? new Date(drive.createdTime) : '',
      '',
      '',
      '',
      `${String(index + 1).padStart(2, '0')}_${sanitizeSheetName(drive.name)}`,
      members.organizers.join(', '),
      members.fileOrganizers.join(', '),
      members.writers.join(', '),
      members.commenters.join(', '),
      members.readers.join(', '),
      driveExternalSharing[drive.id] || 'なし',
      domainUsersOnly,
      driveMembersOnly,
      sharingFoldersRequiresOrganizer,
      copyRequiresWriter,
      '未処理',
      `https://drive.google.com/drive/folders/${drive.id}`
    ];
  });

  if (data.length > 0) {
    sheet.getRange(SHEET_COLUMNS.DATA_ROW_START, 1, data.length, headers.length).setValues(data);
  }

  sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function updateMasterSheetRow(masterSheet, row, stats) {
  masterSheet.getRange(row, SHEET_COLUMNS.MASTER.FILE_COUNT).setValue(stats.totalFiles);

  const sizeInGB = (stats.totalSize / (1024 * 1024 * 1024)).toFixed(2);
  masterSheet.getRange(row, SHEET_COLUMNS.MASTER.SIZE_GB).setValue(parseFloat(sizeInGB));

  const currentExternalStatus = masterSheet.getRange(row, SHEET_COLUMNS.MASTER.EXTERNAL_SHARE).getValue() || '';
  let externalStatus = '';

  const fileExternalCount = stats.externalShareCount || 0;

  if (currentExternalStatus && currentExternalStatus !== 'なし') {
    if (fileExternalCount > 0) {
      externalStatus = `${currentExternalStatus}, ファイルレベル外部共有あり(${fileExternalCount}件)`;
    } else {
      externalStatus = currentExternalStatus;
    }
  } else {
    externalStatus = fileExternalCount > 0 ? `ファイルレベル外部共有あり(${fileExternalCount}件)` : 'なし';
  }

  masterSheet.getRange(row, SHEET_COLUMNS.MASTER.EXTERNAL_SHARE).setValue(externalStatus);
  masterSheet.getRange(row, SHEET_COLUMNS.MASTER.STATUS).setValue('完了');
}

function setupDriveSheetHeaders(sheet) {
  const headers = [
    'レベル', 'パス', '種別', '名前', 'ID', '親ID', '作成者', '作成日',
    '更新日', 'サイズ', '管理者', 'コンテンツ管理者', '投稿者', '編集者', '閲覧者(コメント可)', '閲覧者', '外部共有', 'URL'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#34a853').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function writeDriveDataToSheetAppend(sheet, items, startRow) {
  if (items.length === 0) return;

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

  for (let i = 0; i < data.length; i += CONFIG.BATCH_SIZE) {
    const batch = data.slice(i, i + CONFIG.BATCH_SIZE);
    const currentStartRow = startRow + i;
    sheet.getRange(currentStartRow, 1, batch.length, batch[0].length).setValues(batch);
  }
}
