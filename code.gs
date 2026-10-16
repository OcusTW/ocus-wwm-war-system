/**
 * 《燕雲十六聲》百業戰戰術指引教學系統 — Google Apps Script (Code.gs)
 * 搭配 Google Sheet 自動建立「成員名單」與「時間軸戰術資料」分頁
 * 支援 initDatabase() 手動/自動建表，以及 doGet / doPost 雲端雙向同步
 *
 * 已部署之 Web App URL (供前端 index.html 常數 GAS_WEB_APP_URL 綁定使用):
 * https://script.google.com/macros/s/AKfycbyH3QWcwyTi3DIv3ArSP7ZW5YlN5-wpCBBj1zBrzdwkc6IFib6Yrp1ZjLgwpEUVTPdjAg/exec
 */

const DEPLOYED_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbyH3QWcwyTi3DIv3ArSP7ZW5YlN5-wpCBBj1zBrzdwkc6IFib6Yrp1ZjLgwpEUVTPdjAg/exec';
const SHEET_MEMBERS = '成員名單';
const SHEET_TIMELINE = '時間軸戰術資料';

/**
 * 1. 初始化資料庫 (可於 GAS 編輯器直接點擊「執行」，亦會在 doGet / doPost 時自動防呆觸發)
 */
function initDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 建立或檢查「成員名單」工作表 ---
  let memberSheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!memberSheet) {
    memberSheet = ss.insertSheet(SHEET_MEMBERS);
  }
  if (memberSheet.getLastRow() === 0) {
    const memberHeaders = ['memberId', 'martialArt1', 'martialArt2', 'group', 'updatedAt'];
    memberSheet.getRange(1, 1, 1, memberHeaders.length).setValues([memberHeaders]);
    memberSheet.getRange(1, 1, 1, memberHeaders.length)
      .setBackground('#1e293b')
      .setFontColor('#f8fafc')
      .setFontWeight('bold');
    memberSheet.setFrozenRows(1);

    const now = new Date().toISOString();
    const defaultMembers = [
      ['蒼狼破軍', '嗟夫刀法', '八方風雷槍', '進攻組', now],
      ['劍影無痕', '積矩九劍', '九曲驚神槍', '進攻組', now],
      ['赤焰驚鴻', '無名劍法', '無名槍法', '進攻組', now],
      ['青嵐聽雨', '明川藥典', '千香引魂蠱', '進攻組', now],
      ['墨染江山', '九重春色', '青山執筆', '防守組', now],
      ['玄鐵重盾', '嗟夫刀法', '無名槍法', '防守組', now],
      ['雲裳花容', '千香引魂蠱', '明川藥典', '防守組', now],
      ['疾風殘影', '泥犁三垢', '粟子游塵', '機動組', now],
      ['夜行千里', '泥犁三垢', '無名劍法', '機動組', now],
      ['追風逐月', '粟子游塵', '積矩九劍', '機動組', now]
    ];
    memberSheet.getRange(2, 1, defaultMembers.length, 5).setValues(defaultMembers);
  }

  // --- 建立或檢查「時間軸戰術資料」工作表 ---
  let timelineSheet = ss.getSheetByName(SHEET_TIMELINE);
  if (!timelineSheet) {
    timelineSheet = ss.insertSheet(SHEET_TIMELINE);
  }
  if (timelineSheet.getLastRow() === 0) {
    const timelineHeaders = [
      'minute',
      'tacticName',
      'commanderSkills',
      'tacticAction',
      'hpStatus',
      'deployments',
      'updatedAt'
    ];
    timelineSheet.getRange(1, 1, 1, timelineHeaders.length).setValues([timelineHeaders]);
    timelineSheet.getRange(1, 1, 1, timelineHeaders.length)
      .setBackground('#1e293b')
      .setFontColor('#f8fafc')
      .setFontWeight('bold');
    timelineSheet.setFrozenRows(1);

    const defaultHp = JSON.stringify({
      blue_base: 100,
      red_base: 100,
      blue_wall_top: 100,
      blue_wall_mid: 100,
      blue_wall_bot: 100,
      red_wall_top: 100,
      red_wall_mid: 100,
      red_wall_bot: 100
    });

    const rows = [];
    const now = new Date().toISOString();
    for (let m = 30; m >= 0; m--) {
      let tacticName = '開局清野與三路推進';
      let skills = ['全隊增益'];
      let action = '機動組迅速清空八個野區累積百業幣，進攻組推進敵方人牆，防守組駐守己方三路人牆。';

      if (m <= 25 && m >= 21) {
        tacticName = '首領爭奪與外野壓制';
        skills = ['全隊增益', '快速復活'];
        action = '主力集結爭奪上／下路野怪BOSS（張豹），機動組同步掃蕩八個野區拉開百業幣差距。';
      } else if (m <= 20 && m >= 16) {
        tacticName = '中場單挑與破牆攻勢';
        skills = ['全隊增益', '打鵝傷害'];
        action = '指派單挑高手進入中場表演場地奪取扣血優勢，其餘小隊清空八個野區並破除敵方人牆。';
      } else if (m <= 15 && m >= 11) {
        tacticName = '中期首領與斬鵝行動';
        skills = ['全隊增益', '打鵝傷害', '快速復活'];
        action = '集結搶奪上／下路野怪BOSS，清剿八個野區，進攻組全力集火敵方大鵝。';
      } else if (m <= 10 && m >= 6) {
        tacticName = '後期野區控制與搬樹推進';
        skills = ['搬樹加速', '無視攔截'];
        action = '掃蕩八個野區補充指揮官百業幣，進攻組與防守組協同護送發財樹並攔截敵方反撲。';
      } else if (m <= 5) {
        tacticName = '決戰首領與護樹衝門';
        skills = ['全隊增益', '搬樹加速', '無視攔截', '快速復活'];
        action = '爭奪最後首領（朱骨）增傷Buff，指揮技能全開，全力護送發財樹穿過己方大門奪勝。';
      }

      rows.push([
        m,
        tacticName,
        JSON.stringify(skills),
        action,
        defaultHp,
        JSON.stringify({}),
        now
      ]);
    }
    timelineSheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }

  return { status: 'ok', message: 'Database initialized successfully', url: DEPLOYED_WEB_APP_URL };
}

/**
 * 2. 讀取資料 API (doGet)
 * 若工作表尚未建立，會自動先執行 initDatabase()
 */
function doGet(e) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 讀取成員名單
    const memberSheet = ss.getSheetByName(SHEET_MEMBERS);
    const memberValues = memberSheet.getDataRange().getValues();
    const members = [];
    for (let i = 1; i < memberValues.length; i++) {
      const row = memberValues[i];
      if (!row[0]) continue;
      members.push({
        memberId: String(row[0]).trim(),
        martialArt1: String(row[1] || '無名劍法').trim(),
        martialArt2: String(row[2] || '無名槍法').trim(),
        group: String(row[3] || '進攻組').trim(),
        updatedAt: row[4] || ''
      });
    }

    // 讀取時間軸戰術資料
    const timelineSheet = ss.getSheetByName(SHEET_TIMELINE);
    const timelineValues = timelineSheet.getDataRange().getValues();
    const timeline = {};
    for (let i = 1; i < timelineValues.length; i++) {
      const row = timelineValues[i];
      const minuteKey = String(row[0]).trim();
      if (minuteKey === '') continue;
      timeline[minuteKey] = {
        minute: Number(minuteKey),
        tacticName: String(row[1] || ''),
        commanderSkills: safeJsonParse(row[2], []),
        tacticAction: String(row[3] || ''),
        hpStatus: safeJsonParse(row[4], {}),
        deployments: safeJsonParse(row[5], {}),
        updatedAt: row[6] || ''
      };
    }

    return buildJsonResponse({
      status: 'ok',
      members: members,
      timeline: timeline
    });
  } catch (err) {
    return buildJsonResponse({
      status: 'error',
      message: err.toString()
    });
  }
}

/**
 * 3. 寫入與更新資料 API (doPost)
 * 採用 text/plain JSON 解析，完美相容 Google 協作平台 iframe 跨網域請求
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawBody = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const payload = JSON.parse(rawBody);
    const action = payload.action || 'saveAll';
    const now = new Date().toISOString();

    // 儲存成員名單 (強制確保同一位 memberId 只出現在一組)
    if ((action === 'saveAll' || action === 'saveMembers') && Array.isArray(payload.members)) {
      const memberSheet = ss.getSheetByName(SHEET_MEMBERS);
      const lastRow = memberSheet.getLastRow();
      if (lastRow > 1) {
        memberSheet.getRange(2, 1, lastRow - 1, 5).clearContent();
      }

      const uniqueMap = new Map();
      payload.members.forEach(function(m) {
        if (m && m.memberId) {
          const id = String(m.memberId).trim();
          if (id) {
            uniqueMap.set(id, [
              id,
              String(m.martialArt1 || '無名劍法'),
              String(m.martialArt2 || '無名槍法'),
              String(m.group || '進攻組'),
              now
            ]);
          }
        }
      });

      const memberRows = Array.from(uniqueMap.values());
      if (memberRows.length > 0) {
        memberSheet.getRange(2, 1, memberRows.length, 5).setValues(memberRows);
      }
    }

    // 儲存時間軸戰術與點位配置資料
    if ((action === 'saveAll' || action === 'saveTimeline') && payload.timeline) {
      const timelineSheet = ss.getSheetByName(SHEET_TIMELINE);
      const dataRange = timelineSheet.getDataRange().getValues();
      const rowMap = {};
      for (let i = 1; i < dataRange.length; i++) {
        rowMap[String(dataRange[i][0]).trim()] = i + 1;
      }

      const incomingTimeline = payload.timeline;
      Object.keys(incomingTimeline).forEach(function(minKey) {
        const item = incomingTimeline[minKey];
        const rowIndex = rowMap[String(minKey)];
        const rowValues = [[
          Number(minKey),
          String(item.tacticName || ''),
          JSON.stringify(item.commanderSkills || []),
          String(item.tacticAction || ''),
          JSON.stringify(item.hpStatus || {}),
          JSON.stringify(item.deployments || {}),
          now
        ]];
        if (rowIndex) {
          timelineSheet.getRange(rowIndex, 1, 1, 7).setValues(rowValues);
        } else {
          timelineSheet.appendRow(rowValues[0]);
        }
      });
    }

    return buildJsonResponse({
      status: 'ok',
      updatedAt: now
    });
  } catch (err) {
    return buildJsonResponse({
      status: 'error',
      message: err.toString()
    });
  } finally {
    lock.releaseLock();
  }
}

function safeJsonParse(str, fallback) {
  if (!str) return fallback;
  try {
    return JSON.parse(str);
  } catch (e) {
    return fallback;
  }
}

function buildJsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}