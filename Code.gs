const API_TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiIsImtpZCI6IjI4YTMxOGY3LTAwMDAtYTFlYi03ZmExLTJjNzQzM2M2Y2NhNSJ9.eyJpc3MiOiJzdXBlcmNlbGwiLCJhdWQiOiJzdXBlcmNlbGw6Z2FtZWFwaSIsImp0aSI6ImY2OWYyMWI0LTI0ZDctNDU3YS1iNmVlLWQwZmYyMzFiODBmNCIsImlhdCI6MTc3MzU0MTMzOCwic3ViIjoiZGV2ZWxvcGVyLzMwZWUzNGRjLTQ3MWItYTI0Mi0yMzdkLTQxZjQ4M2YwY2I3YSIsInNjb3BlcyI6WyJjbGFzaCJdLCJsaW1pdHMiOlt7InRpZXIiOiJkZXZlbG9wZXIvc2lsdmVyIiwidHlwZSI6InRocm90dGxpbmcifSx7ImNpZHJzIjpbIjQ1Ljc5LjIxOC43OSJdLCJ0eXBlIjoiY2xpZW50In1dfQ.xS-vUuYz1uPrfl0Funi8FUTID9Nvl_fcpp6lALVjgtUKQTjgI6Z-ydZQ7MEL_Xne5YS-xvSW0X24nVenZvluww";
const CLAN_TAG = "%23YQP0GJ9";
const BASE_URL = "https://cocproxy.royaleapi.dev/v1";

// Minimum expected hero totals by TH level (for rush detection)
const MIN_HERO_TOTALS = { 9: 20, 10: 50, 11: 100, 12: 150, 13: 200, 14: 260, 15: 330, 16: 400, 17: 465 };

// Thresholds for advanced analytics flags
const LIABILITY_STARS_MULTIPLIER = 2.5;   // avg stars/defense above this = high liability
const MIN_TH_FOR_DONATION_CHECK = 10;     // TH level below which donations aren't flagged
const MIN_EXPECTED_DONATIONS = 100;       // donations below this at eligible TH = low donor
const MAX_INACTIVITY_DAYS = 999;          // sentinel value when last-seen timestamp is unavailable
const INACTIVITY_THRESHOLD_DAYS = 14;    // days without activity before player is flagged
const TH_MISMATCH_THRESHOLD = 2;         // TH level difference that triggers mismatch alert

// --- Helper: DRY content-service builder ---
function buildJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- Helper: centralized API fetch with error handling ---
function fetchApi(url, options) {
  try {
    const res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() !== 200) return null;
    return JSON.parse(res.getContentText());
  } catch(e) {
    console.warn("fetchApi error for " + url + ": " + e);
    return null;
  }
}

// --- Helper: compute rush score (% of expected hero levels missing for this TH) ---
function calcRushScore(heroLevels, thLevel) {
  const heroTotal = Object.values(heroLevels).reduce((s, v) => s + (parseInt(v) || 0), 0);
  const expected = MIN_HERO_TOTALS[thLevel] || 0;
  if (expected === 0) return 0;
  return Math.max(0, Math.round(((expected - heroTotal) / expected) * 100));
}

// Shared helper: returns true when a defender absorbs too many stars relative to attacks received
function isHighLiabilityDef(defenses, defStars) {
  return defenses >= 2 && defStars >= Math.floor(defenses * LIABILITY_STARS_MULTIPLIER);
}

// Function to programmatically create the 30-minute auto-update trigger
function createAutoSyncTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'syncClanData') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    ScriptApp.newTrigger('syncClanData')
      .timeBased()
      .everyMinutes(30)
      .create();
    console.log("Auto-sync trigger successfully created for every 30 minutes.");
  } catch(e) {
    console.error("createAutoSyncTrigger failed: " + e);
  }
}

// The API Router
function doGet(e) {
  if (!e || !e.parameter || !e.parameter.action) {
    return ContentService.createTextOutput(JSON.stringify({ error: "No action specified." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = e.parameter.action;
  let responseData = {};

  try {
    if (action === 'getDashboardData') {
      responseData = getDashboardData();
    } else if (action === 'syncClanData') {
      syncClanData();
      responseData = { status: "success", message: "Clan data synced successfully." };
    } else {
      responseData = { error: "Unknown action requested." };
    }
  } catch (error) {
    responseData = { error: error.message };
  }

  return buildJsonResponse(responseData);
}

function cleanBuggedData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const compIdx = data[0].indexOf("Completion %");
  const heroIdx = data[0].indexOf("Hero Total");
  if (compIdx === -1 || heroIdx === -1) return;

  const goodRows = [data[0]];
  let deletedCount = 0;
  for (let i = 1; i < data.length; i++) {
    let pct = parseFloat(data[i][compIdx]);
    let heroStr = String(data[i][heroIdx]);
    if (pct > 100 || heroStr.includes("/305")) {
      deletedCount++;
    } else {
      goodRows.push(data[i]);
    }
  }
  if (deletedCount > 0) {
    sheet.clearContents();
    sheet.getRange(1, 1, goodRows.length, goodRows[0].length).setValues(goodRows);
  }
  console.log(`Successfully cleaned ${deletedCount} corrupted rows.`);
}

function setupMollyMetrics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Data")) ss.getActiveSheet().setName("Data");
  if (!ss.getSheetByName("Activity_Logs")) ss.insertSheet("Activity_Logs");
  if (!ss.getSheetByName("Capital_Raids")) ss.insertSheet("Capital_Raids");
  if (!ss.getSheetByName("War_Data")) ss.insertSheet("War_Data");
  if (!ss.getSheetByName("CWL_History")) ss.insertSheet("CWL_History");
  if (!ss.getSheetByName("Activity_Timeline")) ss.insertSheet("Activity_Timeline");
}

function syncClanData() {
  if (API_TOKEN === "PASTE_YOUR_API_TOKEN_HERE") throw new Error("Please paste your API Token.");

  const options = { "method": "GET", "headers": { "Authorization": "Bearer " + API_TOKEN }, "muteHttpExceptions": true };

  const clanInfo = fetchApi(`${BASE_URL}/clans/${CLAN_TAG}`, options);
  if (!clanInfo || !clanInfo.memberList) return;
  const membersData = clanInfo.memberList;

  try { fetchCapitalRaids(options); } catch(e) { console.error("Capital raids sync failed: " + e); }
  try { fetchWarData(options); } catch(e) { console.error("War data sync failed: " + e); }
  try { updateCwlDatabase(options); } catch(e) { console.error("CWL database sync failed: " + e); }

  const requests = membersData.map(m => ({
    url: `${BASE_URL}/players/${m.tag.replace("#", "%23")}`,
    headers: { "Authorization": "Bearer " + API_TOKEN },
    muteHttpExceptions: true
  }));

  const responses = UrlFetchApp.fetchAll(requests);
  const now = new Date();

  const activitySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activity_Logs");
  const activitySheetData = activitySheet.getDataRange().getValues();
  const activityTagIndex = new Map();
  for (let i = 1; i < activitySheetData.length; i++) {
    if (activitySheetData[i][0]) activityTagIndex.set(activitySheetData[i][0], {
      rowIndex: i + 1,
      lastSeen: activitySheetData[i][2], donations: activitySheetData[i][3],
      received: activitySheetData[i][4], attacks: activitySheetData[i][5],
      score: activitySheetData[i][6] || 0
    });
  }

  const allPlayerData = [];
  responses.forEach(res => {
    try {
      if (res.getResponseCode() === 200) {
        const playerData = JSON.parse(res.getContentText());
        const extracted = extractPlayerData(playerData, now);
        allPlayerData.push({ tag: playerData.tag, extracted });
        updateActivityLog(playerData, now, activitySheet, activityTagIndex);
      }
    } catch(e) {
      console.error("Player processing error: " + e);
    }
  });

  if (allPlayerData.length > 0) batchWriteToDataSheet(allPlayerData);
  CacheService.getScriptCache().remove("mollyDashboardData");
}

function extractPlayerData(data, now) {
  let extractedLevels = {};
  let currentTotal = 0, heroTotal = 0, petTotal = 0, labTotal = 0;

  const heroMap = { "Barbarian King": 105, "Archer Queen": 105, "Grand Warden": 80, "Royal Champion": 55, "Minion Prince": 95, "Dragon Duke": 25 };
  if (data.heroes) {
    data.heroes.forEach(hero => {
      if (hero.village === "home" && heroMap[hero.name]) { currentTotal += hero.level; heroTotal += hero.level; extractedLevels[hero.name] = hero.level; }
    });
  }

  const petMap = { "Electro Owl": 15, "Unicorn": 15, "Frosty": 15, "Diggy": 10, "Phoenix": 10, "Spirit Fox": 10, "Angry Jelly": 10, "Sneezy": 10, "Greedy Raven": 10 };
  if (data.troops) {
    data.troops.forEach(t => {
      if (t.village === "home") {
        if (petMap[t.name]) { currentTotal += t.level; petTotal += t.level; extractedLevels[t.name] = t.level; } else { labTotal += t.level; }
      }
    });
  }
  if (data.spells) data.spells.forEach(s => { if(s.village === "home") labTotal += s.level; });

  let equipmentList = [];
  if (data.heroEquipment) {
    equipmentList = data.heroEquipment.map(eq => ({ name: eq.name, level: eq.level }));
  }

  extractedLevels["Timestamp"] = now.getTime();
  extractedLevels["Username"] = data.name;
  extractedLevels["TH"] = data.townHallLevel || 0;
  extractedLevels["Lab Total"] = labTotal;
  extractedLevels["Completion %"] = ((currentTotal / 570) * 100).toFixed(2);
  extractedLevels["Hero Total"] = `${heroTotal}/465`;
  extractedLevels["Pet Total"] = `${petTotal}/105`;
  extractedLevels["Donations"] = data.donations || 0;
  extractedLevels["Donations Received"] = data.donationsReceived || 0;
  extractedLevels["War Stars"] = data.warStars || 0;
  extractedLevels["Equipment"] = JSON.stringify(equipmentList);
  return extractedLevels;
}

function batchWriteToDataSheet(allPlayerData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues()[0];
  const requiredHeaders = ["Tag", "Username", "Timestamp", "TH", "Completion %", "Hero Total", "Pet Total", "Lab Total", "Barbarian King", "Archer Queen", "Grand Warden", "Royal Champion", "Minion Prince", "Dragon Duke", "Electro Owl", "Unicorn", "Frosty", "Diggy", "Phoenix", "Spirit Fox", "Angry Jelly", "Sneezy", "Greedy Raven", "Donations", "Donations Received", "War Stars", "Equipment"];
  if (!headers[0]) headers = [];
  requiredHeaders.forEach(req => { if (headers.indexOf(req) === -1) { headers.push(req); sheet.getRange(1, headers.length).setValue(req); } });

  const newRows = allPlayerData.map(({ tag, extracted }) =>
    headers.map((h, i) => i === 0 ? tag : (extracted[h] !== undefined ? extracted[h] : ""))
  );
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
}

function updateDataSheet(tag, extractedLevels) {
  batchWriteToDataSheet([{ tag, extracted: extractedLevels }]);
}

function updateActivityLog(playerData, now, sheet, activityTagIndex) {
  const timelineSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activity_Timeline");
  const tag = playerData.tag;
  const currentStats = { donations: playerData.donations || 0, received: playerData.donationsReceived || 0, attacks: playerData.attackWins || 0 };

  const existing = activityTagIndex.get(tag);
  if (!existing) {
    sheet.appendRow([tag, playerData.name, now.getTime(), currentStats.donations, currentStats.received, currentStats.attacks, 1]);
    activityTagIndex.set(tag, { rowIndex: sheet.getLastRow(), lastSeen: now.getTime(), donations: currentStats.donations, received: currentStats.received, attacks: currentStats.attacks, score: 1 });
    if(timelineSheet) timelineSheet.appendRow([now.getTime()]);
  } else {
    const prev = { lastSeen: existing.lastSeen, donations: existing.donations, received: existing.received, attacks: existing.attacks, score: existing.score };
    let hasChanged = (currentStats.donations !== prev.donations) || (currentStats.received !== prev.received) || (currentStats.attacks !== prev.attacks);
    let newLastSeen = hasChanged ? now.getTime() : prev.lastSeen; let newScore = hasChanged ? prev.score + 1 : prev.score;
    sheet.getRange(existing.rowIndex, 1, 1, 7).setValues([[tag, playerData.name, newLastSeen, currentStats.donations, currentStats.received, currentStats.attacks, newScore]]);

    if (hasChanged && timelineSheet) timelineSheet.appendRow([now.getTime()]);
  }
}

function fetchCapitalRaids(options) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Capital_Raids");
  const data = fetchApi(`${BASE_URL}/clans/${CLAN_TAG}/capitalraidseasons`, options);
  if (!data || !data.items || data.items.length === 0) return;

  sheet.clear(); sheet.appendRow(["Tag", "Name", "Attacks", "AttackLimit", "BonusAttacks", "Looted", "HISTORY_JSON"]);
  let historyData = data.items.map(season => {
    let d = season.startTime.substring(0,8); return { date: d.substring(4,6) + "/" + d.substring(6,8) + "/" + d.substring(2,4), loot: season.capitalTotalLoot };
  }).reverse();
  sheet.getRange(1, 7).setValue(JSON.stringify(historyData));

  (data.items[0].members || []).forEach(m => { sheet.appendRow([m.tag, m.name, m.attacks, m.attackLimit, m.bonusAttackLimit, m.capitalResourcesLooted]); });
}

function updateCwlDatabase(options) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CWL_History");
  if (!sheet) return;
  const group = fetchApi(`${BASE_URL}/clans/${CLAN_TAG}/currentwar/leaguegroup`, options);
  if (!group || group.state === "notInWar" || !group.season) return;
  let dataRange = sheet.getDataRange().getValues(); let rowIndex = -1;
  for (let i = 0; i < dataRange.length; i++) { if (dataRange[i][0] === group.season) { rowIndex = i + 1; break; } }
  let payload = JSON.stringify({ season: group.season, state: group.state });
  if (rowIndex === -1) { sheet.appendRow([group.season, payload]); } else { sheet.getRange(rowIndex, 2).setValue(payload); }
}

// THE ESPIONAGE ENGINE: Deep CWL Analytics & Target Patterns
function fetchWarData(options) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("War_Data");
  sheet.clear();
  let currentWarInfo = { type: "NONE", data: null, cwlDetails: [], myCwlStats: {} };

  const group = fetchApi(`${BASE_URL}/clans/${CLAN_TAG}/currentwar/leaguegroup`, options);
  if (group && group.state !== "notInWar") {
    currentWarInfo.type = "CWL";
    currentWarInfo.data = { season: group.season, state: group.state };
    let cwlWars = [];
    let myClanTag = decodeURIComponent(CLAN_TAG);
    let enemyAttackPatterns = {};
    let myCwlStats = {};

    if (group.rounds) {
      // Batch-fetch all CWL war data in parallel instead of sequential calls
      const warRequests = [];
      for (let round of group.rounds) {
        for (let wTag of (round.warTags || [])) {
          if (wTag !== "#0") {
            warRequests.push({
              url: `${BASE_URL}/clanwarleagues/wars/${wTag.replace("#", "%23")}`,
              headers: { "Authorization": "Bearer " + API_TOKEN },
              muteHttpExceptions: true
            });
          }
        }
      }

      const warResponses = warRequests.length > 0 ? UrlFetchApp.fetchAll(warRequests) : [];
      warResponses.forEach(wRes => {
        if (wRes.getResponseCode() !== 200) return;
        let wData;
        try { wData = JSON.parse(wRes.getContentText()); } catch(e) { return; }

        [wData.clan, wData.opponent].forEach(clan => {
          if (!clan) return;
          if (!enemyAttackPatterns[clan.tag]) {
            enemyAttackPatterns[clan.tag] = { mirrors: 0, dips: 0, reaches: 0, total: 0, threeStars: 0, extremeDips: 0, extremeReaches: 0 };
          }
          let p = enemyAttackPatterns[clan.tag];
          let oppClan = clan.tag === wData.clan.tag ? wData.opponent : wData.clan;
          const oppMemberMap = new Map((oppClan.members || []).map(d => [d.tag, d]));

          (clan.members || []).forEach(m => {
            if (!m.attacks) return;
            m.attacks.forEach(atk => {
              let defender = oppMemberMap.get(atk.defenderTag);
              if (!defender) return;
              p.total++;
              if (atk.stars === 3) p.threeStars++;
              if (m.mapPosition === defender.mapPosition) p.mirrors++;
              else if (m.mapPosition < defender.mapPosition) p.dips++;
              else p.reaches++;
              // TH mismatch detection: extreme = TH_MISMATCH_THRESHOLD+ TH level difference
              const thDiff = (m.townhallLevel || 0) - (defender.townhallLevel || 0);
              if (thDiff >= TH_MISMATCH_THRESHOLD) p.extremeDips++;
              else if (thDiff <= -TH_MISMATCH_THRESHOLD) p.extremeReaches++;
            });
          });
        });

        if (wData.clan.tag === myClanTag || wData.opponent.tag === myClanTag) {
          let enemyClan = wData.clan.tag === myClanTag ? wData.opponent : wData.clan;
          let myClan = wData.clan.tag === myClanTag ? wData.clan : wData.opponent;

          (myClan.members || []).forEach(m => {
            if (!myCwlStats[m.tag]) myCwlStats[m.tag] = { attacks: 0, threeStars: 0, defenses: 0, defStars: 0 };
            if (m.attacks) m.attacks.forEach(atk => {
              myCwlStats[m.tag].attacks++;
              if (atk.stars === 3) myCwlStats[m.tag].threeStars++;
            });
          });

          (enemyClan.members || []).forEach(m => {
            if (!m.attacks) return;
            m.attacks.forEach(atk => {
              if (!myCwlStats[atk.defenderTag]) myCwlStats[atk.defenderTag] = { attacks: 0, threeStars: 0, defenses: 0, defStars: 0 };
              myCwlStats[atk.defenderTag].defenses++;
              myCwlStats[atk.defenderTag].defStars += atk.stars;
            });
          });

          // Build enemy lineup with per-member high-liability detection
          const attacksReceived = {};
          (myClan.members || []).forEach(m => {
            if (!m.attacks) return;
            m.attacks.forEach(atk => {
              if (!attacksReceived[atk.defenderTag]) attacksReceived[atk.defenderTag] = { defenses: 0, defStars: 0 };
              attacksReceived[atk.defenderTag].defenses++;
              attacksReceived[atk.defenderTag].defStars += atk.stars;
            });
          });

          let enemyLineup = (enemyClan.members || []).map(m => {
            let ds = attacksReceived[m.tag] || { defenses: 0, defStars: 0 };
            return {
              name: m.name, tag: m.tag, th: m.townhallLevel, mapPosition: m.mapPosition,
              attacks: m.attacks ? m.attacks.length : 0,
              isHighLiability: isHighLiabilityDef(ds.defenses, ds.defStars)
            };
          }).sort((a,b) => a.mapPosition - b.mapPosition);

          let myLineup = (myClan.members || []).map(m => ({
            name: m.name, tag: m.tag, th: m.townhallLevel, mapPosition: m.mapPosition
          })).sort((a,b) => a.mapPosition - b.mapPosition);

          cwlWars.push({ state: wData.state, opponent: { name: enemyClan.name, tag: enemyClan.tag }, enemyLineup, myLineup });
        }
      });
    }

    // Attach attack patterns and compute true hit rates
    cwlWars.forEach(w => {
      let pat = enemyAttackPatterns[w.opponent.tag] || { mirrors: 0, dips: 0, reaches: 0, total: 0, threeStars: 0, extremeDips: 0, extremeReaches: 0 };
      pat.hitRate = pat.total > 0 ? Math.round((pat.threeStars / pat.total) * 100) : 0;
      w.attackPattern = pat;
    });

    // Compute per-player CWL hit rate and liability flag
    for (let tag in myCwlStats) {
      let s = myCwlStats[tag];
      s.hitRate = s.attacks > 0 ? Math.round((s.threeStars / s.attacks) * 100) : 0;
      s.isHighLiability = isHighLiabilityDef(s.defenses, s.defStars);
    }

    currentWarInfo.cwlDetails = cwlWars;
    currentWarInfo.myCwlStats = myCwlStats;
  }

  if (currentWarInfo.type === "NONE") {
    const war = fetchApi(`${BASE_URL}/clans/${CLAN_TAG}/currentwar`, options);
    if (war && war.state !== "notInWar") {
      currentWarInfo.type = "REGULAR";
      currentWarInfo.data = { state: war.state, opponent: { name: war.opponent ? war.opponent.name : "Unknown" } };
    }
  }

  let jsonPayload = JSON.stringify(currentWarInfo);
  if (jsonPayload.length > 49000) {
    currentWarInfo.cwlDetails = currentWarInfo.cwlDetails.slice(-3);
    jsonPayload = JSON.stringify(currentWarInfo);
  }
  sheet.getRange(1, 1).setValue(jsonPayload);
}

function getDashboardData() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("mollyDashboardData");
  if (cachedData) return JSON.parse(cachedData);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Data").getDataRange().getValues();
  if (!dataSheet || dataSheet.length === 0 || !dataSheet[0] || dataSheet[0].length === 0) {
    return { players: [], clanPct: 0, clanHeroes: "0/0", clanPets: "0/0", clanDonations: 0, clanReceived: 0, clanHistoryDates: [], clanHistoryPcts: [], clanCapitalLoot: 0, capitalHistory: [], cwlHistory: [], warInfo: { type: "NONE", data: null, cwlDetails: [], myCwlStats: {} }, activityGraphs: { hourly: [], daily: [] } };
  }

  const activityData = ss.getSheetByName("Activity_Logs").getDataRange().getValues();
  const capitalData = ss.getSheetByName("Capital_Raids").getDataRange().getValues();
  let cwlHistoryArray = [];
  try {
    let cwlDbData = ss.getSheetByName("CWL_History").getDataRange().getValues();
    for (let i = 0; i < cwlDbData.length; i++) {
      if (cwlDbData[i][1]) {
        try { cwlHistoryArray.push(JSON.parse(cwlDbData[i][1])); } catch(parseErr) { console.warn(`Skipping malformed CWL record at row ${i + 1}`); }
      }
    }
  } catch(e) {}
  cwlHistoryArray.reverse();

  let hourlyActivity = new Array(24).fill(0);
  let dailyActivity = new Array(7).fill(0);
  try {
    let timelineSheet = ss.getSheetByName("Activity_Timeline");
    if (timelineSheet) {
      let timelineData = timelineSheet.getDataRange().getValues();
      const tz = "America/Chicago";
      // Starts at i=0 and skips invalid rows to prevent breaking the chart logic
      for (let i = 0; i < timelineData.length; i++) { 
        let ts = timelineData[i][0];
        if (!ts || isNaN(new Date(ts).getTime())) continue; 

        let d = new Date(ts);
        let hour = parseInt(Utilities.formatDate(d, tz, "H"));
        let dayRaw = parseInt(Utilities.formatDate(d, tz, "u"));
        let dayIdx = dayRaw === 7 ? 0 : dayRaw;
        hourlyActivity[hour]++;
        dailyActivity[dayIdx]++;
      }
    }
  } catch(e) { 
    console.warn("Timeline parsing error: " + e); 
  }

  let warRaw = ""; try { warRaw = ss.getSheetByName("War_Data").getRange(1, 1).getValue(); } catch(e) {}
  let warParsed = { type: "NONE", data: null, cwlDetails: [], myCwlStats: {} };
  try { if (warRaw) warParsed = JSON.parse(warRaw); } catch(e) {}

  let activityMap = {}; if (activityData && activityData.length > 1) { for(let i=1; i<activityData.length; i++) activityMap[activityData[i][0]] = { lastSeenTs: activityData[i][2], activityScore: activityData[i][6] }; }
  let capitalMap = {}; let totalClanLoot = 0; if (capitalData && capitalData.length > 1) { for(let i=1; i<capitalData.length; i++) { capitalMap[capitalData[i][0]] = { attacks: capitalData[i][2], limit: capitalData[i][3] + capitalData[i][4], looted: capitalData[i][5] }; totalClanLoot += capitalData[i][5]; } }
  let capitalHistory = []; try { let histStr = ss.getSheetByName("Capital_Raids").getRange(1, 7).getValue(); if (histStr) capitalHistory = JSON.parse(histStr); } catch(e) {}

  let playerMap = {};
  const headers = dataSheet[0];
  const headerMap = {};
  headers.forEach((h, i) => { headerMap[h] = i; });
  const tagIdx = 0; const nameIdx = headerMap["Username"] ?? -1; const tsIdx = headerMap["Timestamp"] ?? -1;
  const compIdx = headerMap["Completion %"] ?? -1;

  for (let i = 1; i < dataSheet.length; i++) {
    const row = dataSheet[i]; const tag = row[tagIdx]; if (!tag || tag === "Tag") continue;
    const ts = tsIdx > -1 && row[tsIdx] ? row[tsIdx] : 0;
    const dateLabel = ts > 0 ? Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), "MM/dd/yy") : "";
    if (!playerMap[tag]) playerMap[tag] = { tag: tag, history: [], latestTs: 0 };
    playerMap[tag].history.push({ date: dateLabel, pct: parseFloat(row[compIdx]) || 0, ts: ts, rowIdx: i });
    if (ts >= playerMap[tag].latestTs || playerMap[tag].latestTs === 0) playerMap[tag].latestTs = ts;
  }

  let results = []; let uniqueDatesMap = {};
  const thirtyDaysAgo = new Date().getTime() - (30 * 24 * 60 * 60 * 1000);
  const heroNames = ["Barbarian King", "Archer Queen", "Grand Warden", "Royal Champion", "Minion Prince", "Dragon Duke"];
  const petNames = ["Electro Owl", "Unicorn", "Frosty", "Diggy", "Phoenix", "Spirit Fox", "Angry Jelly", "Sneezy", "Greedy Raven"];

  let clanTotalDonations = 0;
  let clanTotalReceived = 0;

  for (let key in playerMap) {
    let p = playerMap[key]; p.history.sort((a, b) => a.ts - b.ts);
    let latestEntry = p.history[p.history.length - 1]; let prevEntry = p.history.length > 1 ? p.history[p.history.length - 2] : latestEntry;
    let row = dataSheet[latestEntry.rowIdx]; let prevRow = dataSheet[prevEntry.rowIdx];
    p.history.forEach(h => { if (!uniqueDatesMap[h.date] || h.ts > uniqueDatesMap[h.date]) uniqueDatesMap[h.date] = h.ts; });
    let chartHistory = p.history.filter(h => h.ts >= thirtyDaysAgo); if (chartHistory.length === 0 && p.history.length > 0) chartHistory.push(p.history[p.history.length - 1]);
    let act = activityMap[key] || { lastSeenTs: 0, activityScore: 0 }; let cap = capitalMap[key] || { attacks: 0, limit: 6, looted: 0 };

    let donations = headerMap["Donations"] !== undefined ? parseInt(row[headerMap["Donations"]]) || 0 : 0;
    let received = headerMap["Donations Received"] !== undefined ? parseInt(row[headerMap["Donations Received"]]) || 0 : 0;
    clanTotalDonations += donations;
    clanTotalReceived += received;

    let cwlStats = warParsed.myCwlStats[p.tag] || { attacks: 0, threeStars: 0, defenses: 0, defStars: 0 };

    let playerObj = {
      name: nameIdx > -1 && row[nameIdx] ? row[nameIdx] : p.tag, tag: p.tag,
      thLevel: headerMap["TH"] !== undefined ? parseInt(row[headerMap["TH"]]) : 0,
      percentage: parseFloat((row[compIdx] || 0).toString()).toFixed(2), prevPct: parseFloat((prevRow[compIdx] || 0).toString()).toFixed(2),
      heroStats: headerMap["Hero Total"] !== undefined ? row[headerMap["Hero Total"]] : "0/465", petStats: headerMap["Pet Total"] !== undefined ? row[headerMap["Pet Total"]] : "0/105",
      labStats: headerMap["Lab Total"] !== undefined ? row[headerMap["Lab Total"]] : 0,
      donations: donations, donationsReceived: received, warStars: headerMap["War Stars"] !== undefined ? row[headerMap["War Stars"]] : 0,
      lastSeenTs: act.lastSeenTs, activityScore: act.activityScore, capitalAttacks: `${cap.attacks}/${cap.limit}`, capitalLooted: cap.looted,
      cwlStats: cwlStats,
      historyDates: chartHistory.map(h => h.date), historyPcts: chartHistory.map(h => h.pct),
      heroes: {}, pets: {}, upgrades: [], recentlyCompleted: []
    };

    const getLvl = (r, name) => { let idx = headerMap[name] ?? -1; return idx > -1 ? parseInt(r[idx]) || 0 : 0; };

    heroNames.forEach(h => { playerObj.heroes[h] = getLvl(row, h); if(playerObj.heroes[h] > getLvl(prevRow, h)) playerObj.recentlyCompleted.push(`${h} to Lvl ${playerObj.heroes[h]}`); });
    petNames.forEach(pet => { playerObj.pets[pet] = getLvl(row, pet); if(playerObj.pets[pet] > getLvl(prevRow, pet)) playerObj.recentlyCompleted.push(`${pet} to Lvl ${playerObj.pets[pet]}`); });

    // --- Advanced analytics flags ---
    const rushScore = calcRushScore(playerObj.heroes, playerObj.thLevel);
    const inactivityDays = act.lastSeenTs > 0 ? Math.floor((new Date().getTime() - act.lastSeenTs) / (1000 * 60 * 60 * 24)) : MAX_INACTIVITY_DAYS;
    const cwlHitRate = cwlStats.attacks > 0 ? Math.round((cwlStats.threeStars / cwlStats.attacks) * 100) : null;
    playerObj.rushScore = rushScore;
    playerObj.isRushed = rushScore > 30;
    playerObj.inactivityDays = inactivityDays;
    playerObj.isInactive = inactivityDays > INACTIVITY_THRESHOLD_DAYS;
    playerObj.isLowDonor = playerObj.thLevel >= MIN_TH_FOR_DONATION_CHECK && donations < MIN_EXPECTED_DONATIONS;
    playerObj.hitRate = cwlHitRate;
    playerObj.isHighLiability = cwlStats.isHighLiability || isHighLiabilityDef(cwlStats.defenses || 0, cwlStats.defStars || 0);

    results.push(playerObj);
  }

  results.sort((a, b) => b.percentage - a.percentage);
  let clanTotalHeroes = 0, clanMaxHeroes = 0; let clanTotalPets = 0, clanMaxPets = 0;
  results.forEach(p => {
    let heroParts = (p.heroStats || "0/465").split('/'); clanTotalHeroes += parseInt(heroParts[0]) || 0; clanMaxHeroes += parseInt(heroParts[1]) || 465;
    let petParts = (p.petStats || "0/105").split('/'); clanTotalPets += parseInt(petParts[0]) || 0; clanMaxPets += parseInt(petParts[1]) || 105;
  });

  let uniqueDatesArray = Object.keys(uniqueDatesMap).map(date => ({ label: date, ts: uniqueDatesMap[date] })).sort((a, b) => a.ts - b.ts);
  let clanHistoryDates = [], clanHistoryPcts = [];
  uniqueDatesArray.filter(d => d.ts >= thirtyDaysAgo).forEach(d => {
    let totalPctAtDate = 0, activePlayerCount = 0;
    for (let key in playerMap) {
      let validRecords = playerMap[key].history.filter(h => h.ts <= (d.ts + 86400000));
      if (validRecords.length > 0) { totalPctAtDate += validRecords[validRecords.length - 1].pct; activePlayerCount++; }
    }
    if (activePlayerCount > 0) { clanHistoryDates.push(d.label); clanHistoryPcts.push(parseFloat((totalPctAtDate / activePlayerCount).toFixed(2))); }
  });

  const finalResponse = {
    players: results, clanPct: clanHistoryPcts.length > 0 ? clanHistoryPcts[clanHistoryPcts.length - 1] : 0,
    clanHeroes: `${clanTotalHeroes}/${clanMaxHeroes}`, clanPets: `${clanTotalPets}/${clanMaxPets}`,
    clanDonations: clanTotalDonations, clanReceived: clanTotalReceived,
    clanHistoryDates: clanHistoryDates, clanHistoryPcts: clanHistoryPcts,
    clanCapitalLoot: totalClanLoot, capitalHistory: capitalHistory,
    cwlHistory: cwlHistoryArray,
    warInfo: { type: warParsed.type, data: warParsed.data, cwlDetails: warParsed.cwlDetails },
    activityGraphs: { hourly: hourlyActivity, daily: dailyActivity }
  };
  try { CacheService.getScriptCache().put("mollyDashboardData", JSON.stringify(finalResponse), 1800); } catch(e) { console.warn("Cache bypassed due to size."); }
  return finalResponse;
}
