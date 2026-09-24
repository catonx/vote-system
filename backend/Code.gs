/** 
 * Code.gs - QuickVote backend
 * - Sheets used:
 *   - "Contexts"  (context_id, title, options, created_at, ends_at)
 *   - "Votes"     (context_id, token, fingerprint, choice, timestamp)
 *   - optional "Tokens" (Token, Identifier, Used)  => if present, tokens must be valid
 *
 * API:
 * GET  ?action=getContexts
 * GET  ?action=getResults&contextId=...
 * POST body JSON { action:"createContext", title:"..", options:[..] }
 * POST body JSON { action:"vote", contextId:"..", token:"..", choice:"..", fingerprint:".." }
 * POST body JSON { action:"deleteContext", contextId:".." }   (admin)
 *
 * Notes:
 * - Token whitelist enforced only if sheet "Tokens" exists.
 * - Double-vote prevented per (contextId, token).
 /**
 * === SHEET NAMES ===
 */
/**
 * === SHEET NAMES ===
 */
const SHEET_VOTES = "Votes";
const SHEET_CONTEXTS = "Contexts";
const SHEET_TOKENS = "Tokens";

function getContextHeaderMap(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return {};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headers.forEach((header, index) => {
    if (header) headerMap[String(header).trim().toLowerCase()] = index;
  });
  return headerMap;
}

function parseContextEndDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (!value) return null;
  const parsed = new Date(String(value));
  return isNaN(parsed.getTime()) ? null : parsed;
}

function getContextById(sheet, contextId) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  const headerMap = getContextHeaderMap(sheet);
  const idIndex = headerMap.context_id;
  if (idIndex === undefined) return null;
  const rows = sheet.getDataRange().getValues();
  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    if (String(row[idIndex]) === String(contextId)) {
      return { row, rowIndex: rowIndex + 1, headerMap };
    }
  }
  return null;
}

function parseOptions(value) {
  if (Array.isArray(value)) return value;
  if (!value) return [];
  try {
    const parsed = JSON.parse(String(value));
    if (Array.isArray(parsed)) return parsed;
    if (typeof parsed === "string") return JSON.parse(parsed);
  } catch (error) {
    Logger.log("Invalid context options JSON: " + error);
  }
  return [];
}

function getMasterKey() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MasterKeys");
  if (!sheet || sheet.getLastRow() < 2) return "";

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const keyColumn = headers.findIndex(header => String(header).trim().toLowerCase() === "key");
  const column = keyColumn === -1 ? 1 : keyColumn + 1;
  return String(sheet.getRange(2, column).getValue() || "").trim();
}

function isValidAdminKey(adminKey) {
  const masterKey = getMasterKey();
  return Boolean(masterKey && adminKey && String(adminKey).trim() === masterKey);
}

function requireAdminKey(data) {
  if (!isValidAdminKey(data && data.adminKey)) {
    return { success: false, code: "INVALID_ADMIN_KEY", message: "Admin key tidak valid." };
  }
  return null;
}

function getOptionName(option) {
  if (option && typeof option === "object") {
    return String(option.name || option.title || "");
  }
  return String(option || "");
}

function isAttendanceContext(context) {
  return Boolean(context && (
    context.headerMap.attendance !== undefined &&
    (context.row[context.headerMap.attendance] === true ||
      String(context.row[context.headerMap.attendance] || "").trim().toLowerCase() === "true")
  ));
}

function validateOpenSession(contextSheet, contextId) {
  const context = getContextById(contextSheet, contextId);
  if (!context) {
    return { success: false, code: "CONTEXT_NOT_FOUND", message: "Voting context not found." };
  }

  const endsAtIndex = context.headerMap.ends_at;
  const endsAt = endsAtIndex === undefined ? null : parseContextEndDate(context.row[endsAtIndex]);
  if (!endsAt) {
    return { success: false, code: "SESSION_NOT_CONFIGURED", message: "Session end time is not configured." };
  }
  if (endsAt.getTime() <= Date.now()) {
    return { success: false, code: "SESSION_ENDED", message: "Session has ended." };
  }
  return { success: true, context, endsAt };
}

function validateOpenContext(contextSheet, contextId, choice) {
  const context = getContextById(contextSheet, contextId);
  if (!context) {
    return { success: false, code: "CONTEXT_NOT_FOUND", message: "Voting context not found." };
  }

  const endsAtIndex = context.headerMap.ends_at;
  const endsAt = endsAtIndex === undefined ? null : parseContextEndDate(context.row[endsAtIndex]);
  if (!endsAt) {
    return { success: false, code: "SESSION_NOT_CONFIGURED", message: "Voting session end time is not configured." };
  }
  if (endsAt.getTime() <= Date.now()) {
    return { success: false, code: "SESSION_ENDED", message: "Voting session has ended." };
  }

  const optionsIndex = context.headerMap.options;
  const options = optionsIndex === undefined ? [] : parseOptions(context.row[optionsIndex]);
  if (!options.some(option => getOptionName(option) === String(choice))) {
    return { success: false, code: "INVALID_CHOICE", message: "Choice is not available for this context." };
  }
  return { success: true, context, endsAt };
}

/**
 * === UNIVERSAL JSON RESPONSE WITH CORS ===
 */
function sendJSON(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
/**
 * Ensure Votes sheet exists and has header including voters column (col 6)
 */
function ensureVotesSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_VOTES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_VOTES);
    sheet.appendRow(["vote_id","contextId", "token", "choice", "timestamp", "fingerprint", "voters"]);
    return sheet;
  }
  // ensure header has at least 6 columns
  //const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const header = sheet.getRange(1, 1, 1, 7).getValues()[0];
  if (header.length < 7 || header[6] !== "voters") {
    // extend header to 6 cols and set proper names (preserve existing existing headers if present)
    const fixedHeader = [
      header[0] || "vote_id",
      header[1] || "contextId",
      header[2] || "token",
      header[3] || "choice",
      header[4] || "timestamp",
      header[5] || "fingerprint",
      "voters"
    ];
    sheet.getRange(1, 1, 1, 7).setValues([fixedHeader]);
  }
  return sheet;
}
/**
 * === CORE ACTION HANDLER ===
 */
function handleAction(action, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contextSheet = ss.getSheetByName("Contexts");
  const votesSheet = ss.getSheetByName("Votes");
  const tokensSheet = ss.getSheetByName("Tokens");

  switch (action) {
    case "validateAdminKey":
      return isValidAdminKey(data && data.adminKey)
        ? { success: true }
        : { success: false, code: "INVALID_ADMIN_KEY", message: "Admin key tidak valid." };

    case "createContext": {
      const adminError = requireAdminKey(data);
      if (adminError) return adminError;
      const { title, options, endsAt, ends_at, attendance } = data;
      const endDate = parseContextEndDate(endsAt || ends_at);
      if (!title || !endDate || endDate.getTime() <= Date.now()) {
        return { success: false, code: "INVALID_END_TIME", message: "A future ends_at value is required." };
      }
      const isAttendance = String(attendance).trim().toLowerCase() === "true";
      let sheet = ss.getSheetByName(SHEET_CONTEXTS);
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_CONTEXTS);
        sheet.appendRow(["context_id", "title", "options", "created_at", "ends_at", "attendance"]);
      } else {
        const headerMap = getContextHeaderMap(sheet);
        if (headerMap.ends_at === undefined) {
          return { success: false, code: "ENDS_AT_COLUMN_MISSING", message: "Add an ends_at column to the Contexts sheet first." };
        }
        if (isAttendance && headerMap.attendance === undefined) {
          return { success: false, code: "ATTENDANCE_COLUMN_MISSING", message: "Add an attendance column to the Contexts sheet first." };
        }
      }
      const contextId = Utilities.getUuid();
      const headerMap = getContextHeaderMap(sheet);
      const row = new Array(Math.max(sheet.getLastColumn(), 5)).fill("");
      row[headerMap.context_id] = contextId;
      row[headerMap.title] = title;
      row[headerMap.options] = typeof options === "string" ? options : JSON.stringify(options || []);
      row[headerMap.created_at] = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
      row[headerMap.ends_at] = endDate;
      if (headerMap.attendance !== undefined) row[headerMap.attendance] = isAttendance;
      sheet.appendRow(row);
      return { success: true, contextId };
    }

    case "getContexts": {
      const sheet = ss.getSheetByName(SHEET_CONTEXTS);
      if (!sheet) return { success: true, contexts: [] };
      const headerMap = getContextHeaderMap(sheet);
      const rows = sheet.getDataRange().getValues().slice(1);
      const contexts = rows.filter(r => r[headerMap.context_id] !== "").map(r => ({
        id: r[headerMap.context_id],
        title: r[headerMap.title],
        options: parseOptions(r[headerMap.options]),
        created_at: r[headerMap.created_at],
        ends_at: headerMap.ends_at === undefined ? "" : r[headerMap.ends_at],
        attendance: headerMap.attendance !== undefined && (
          r[headerMap.attendance] === true ||
          String(r[headerMap.attendance] || "").trim().toLowerCase() === "true"
        )
      }));
      return { success: true, contexts };
    }

    case "getParticipant": {
      const token = String(data && data.token || "").trim();
      if (!token || !isValidToken(token)) {
        return { success: false, code: "INVALID_TOKEN", message: "Invalid or unauthorized token." };
      }
      const sheet = ss.getSheetByName(SHEET_TOKENS);
      const headerMap = getTokenHeaderMap(sheet);
      const tokenIndex = headerMap.token === undefined ? 0 : headerMap.token;
      const identifierIndex = headerMap.identifier === undefined ? 1 : headerMap.identifier;
      const rows = sheet.getDataRange().getValues().slice(1);
      const row = rows.find(item => String(item[tokenIndex] || "").trim() === token);
      return row
        ? { success: true, identifier: String(row[identifierIndex] || "").trim() }
        : { success: false, code: "INVALID_TOKEN", message: "Invalid or unauthorized token." };
    }

    case "getAttendanceStatus": {
      const contextId = String(data && data.contextId || "").trim();
      const token = String(data && data.token || "").trim();
      if (!contextId || !token || !isValidToken(token)) {
        return { success: false, code: "INVALID_TOKEN", message: "Invalid or unauthorized token." };
      }

      const sheet = ss.getSheetByName(SHEET_VOTES);
      if (!sheet || sheet.getLastRow() < 2) return { success: true, recorded: false };
      const rows = sheet.getDataRange().getValues().slice(1);
      const attendance = rows.find(row => (
        String(row[1]) === contextId &&
        String(row[2]).trim() === token &&
        String(row[3]).trim().toLowerCase() === "hadir"
      ));
      return { success: true, recorded: Boolean(attendance) };
    }

    case "getTokens": {
      const adminError = requireAdminKey(data);
      if (adminError) return adminError;
      const sheet = ss.getSheetByName(SHEET_TOKENS);
      if (!sheet || sheet.getLastRow() < 2) return { success: true, tokens: [] };
      const headerMap = getTokenHeaderMap(sheet);
      const tokenIndex = headerMap.token === undefined ? 0 : headerMap.token;
      const identifierIndex = headerMap.identifier === undefined ? 1 : headerMap.identifier;
      const usedIndex = headerMap.used === undefined ? 2 : headerMap.used;
      const rows = sheet.getDataRange().getValues().slice(1);
      const tokens = rows
        .filter(row => String(row[tokenIndex] || "").trim())
        .map(row => ({
          token: String(row[tokenIndex]).trim(),
          identifier: String(row[identifierIndex] || "").trim(),
          used: row[usedIndex] === true || String(row[usedIndex] || "").trim().toLowerCase() === "true"
        }));
      return { success: true, tokens };
    }

    case "deleteOption": {
      const adminError = requireAdminKey(data);
      if (adminError) return adminError;
      const { contextId, optionName } = data;
      if (!contextId || !optionName) {
        return { success: false, code: "MISSING_OPTION_FIELDS", message: "Context and option are required." };
      }

      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        const context = getContextById(contextSheet, contextId);
        if (!context) {
          return { success: false, code: "CONTEXT_NOT_FOUND", message: "Voting context not found." };
        }

        const optionsIndex = context.headerMap.options;
        const options = optionsIndex === undefined ? [] : parseOptions(context.row[optionsIndex]);
        if (options.length <= 1) {
          return { success: false, code: "MINIMUM_OPTION_REQUIRED", message: "At least one option must remain." };
        }

        const optionIndex = options.findIndex(option => getOptionName(option) === String(optionName));
        if (optionIndex === -1) {
          return { success: false, code: "OPTION_NOT_FOUND", message: "Option not found in this context." };
        }

        options.splice(optionIndex, 1);
        contextSheet.getRange(context.rowIndex, optionsIndex + 1).setValue(JSON.stringify(options));
        return { success: true, message: "Option deleted.", contextId, optionName };
      } finally {
        lock.releaseLock();
      }
    }

    case "submitVote": {
      const { contextId, token, choice, fingerprint } = data;
      if (!contextId || !token || !choice) {
        return { success: false, message: "Missing required vote fields." };
      }

      const contextValidation = validateOpenContext(contextSheet, contextId, choice);
      if (!contextValidation.success) return contextValidation;
      if (isAttendanceContext(contextValidation.context)) {
        return { success: false, code: "ATTENDANCE_ONLY", message: "This context is for attendance only." };
      }

      // If Tokens sheet exists, enforce whitelist
      if (ss.getSheetByName("Tokens") && !isValidToken(token)) {
        return { success: false, code: "INVALID_TOKEN", message: "Invalid or unauthorized token." };
      }

      // Serialize duplicate checking and append so concurrent requests cannot both pass.
      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      let sheet;
      try {
        const finalContextValidation = validateOpenContext(contextSheet, contextId, choice);
        if (!finalContextValidation.success) return finalContextValidation;
        sheet = ensureVotesSheet(ss);

        // check duplicate vote (contextId+token)
        const lastRow = sheet.getLastRow();
        const rows = lastRow > 1
          ? sheet.getRange(2, 1, lastRow - 1, 7).getValues()
          : [];

        const voteId = Utilities.getUuid();
        const already = rows.find(r => String(r[1]) === String(contextId) && String(r[2]) === String(token));
        if (already) {
          return { success: false, code: "TOKEN_ALREADY_USED", message: "Token already used for this context" };
        }

        const timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
        const timestamp = Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd HH:mm:ss");
        sheet.appendRow([voteId, contextId, token, choice, timestamp, fingerprint || "", ""]);

        Logger.log("submitVote return payload: " + JSON.stringify({
          success: true,
          message: "Vote recorded successfully",
          voteId: voteId
        }));
        return { success: true, message: "Vote recorded successfully", voteId: voteId };
      } finally {
        lock.releaseLock();
      }
    }

    case "submitAttendance": {
      const { contextId, token, fingerprint } = data;
      if (!contextId || !token) {
        return { success: false, code: "MISSING_ATTENDANCE_FIELDS", message: "Context and token are required." };
      }

      const sessionValidation = validateOpenSession(contextSheet, contextId);
      if (!sessionValidation.success) return sessionValidation;
      if (!isAttendanceContext(sessionValidation.context)) {
        return { success: false, code: "NOT_ATTENDANCE_CONTEXT", message: "This context is not for attendance." };
      }
      if (ss.getSheetByName(SHEET_TOKENS) && !isValidToken(token)) {
        return { success: false, code: "INVALID_TOKEN", message: "Invalid or unauthorized token." };
      }

      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        const finalValidation = validateOpenSession(contextSheet, contextId);
        if (!finalValidation.success) return finalValidation;
        const sheet = ensureVotesSheet(ss);
        const lastRow = sheet.getLastRow();
        const rows = lastRow > 1
          ? sheet.getRange(2, 1, lastRow - 1, 7).getValues()
          : [];
        const already = rows.find(row => String(row[1]) === String(contextId) && String(row[2]) === String(token));
        if (already) {
          return { success: false, code: "ATTENDANCE_ALREADY_RECORDED", message: "Attendance already recorded for this context." };
        }

        const voteId = Utilities.getUuid();
        const timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
        const timestamp = Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd HH:mm:ss");
        sheet.appendRow([voteId, contextId, token, "Hadir", timestamp, fingerprint || "", ""]);
        return { success: true, message: "Attendance recorded successfully.", voteId };
      } finally {
        lock.releaseLock();
      }
    }

    case "saveVoterInfo": {
      // expects data.row and data.voter (string "Name-Email")
      
      // const row = Number(data.row);
      // const voter = data.voter;
      // if (!row || !voter) {
        // return { success: false, message: "Missing row or voter" };
      // }
      const { voteId, voter } = data;
      const sheet = ensureVotesSheet(ss);
      const rows = sheet.getDataRange().getValues();
      //if (row < 2 || row > sheet.getLastRow()) {
        // return { success: false, message: "Invalid row" };
      // }
      // sheet.getRange(row, 6).setValue(voter);
      // return { success: true, message: "Voter info saved", row };
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === voteId) {
          sheet.getRange(i + 1, 7).setValue(voter); // kolom voters
          return { success: true, message: "Voter info saved", rows };
        }
      }
      return { success: false, message: "Vote not found" };
    }

    case "getResults": {
      const { contextId } = data;
      const sheet = ss.getSheetByName(SHEET_VOTES);
      if (!sheet) return { success: true, results: {} };

      const rows = sheet.getDataRange().getValues().slice(1).filter(r => r[1] === contextId);
      const counts = {};
      for (let r of rows) {
        counts[r[3]] = (counts[r[3]] || 0) + 1;
      }
      return { success: true, results: counts };
    }
    // case "vote": {
      // const { token, contextId, choice } = data;
      // if (!isValidToken(token)) {
      //  return { success: false, message: "Invalid or unauthorized token." };
      // }
      // return recordVote(contextId, token, choice);
    // }
    case "deleteVote": {
      const voteId = data.voteId;
      if (!voteId) {
        return { success: false, error: "voteId is required" };
      }
      // const sheet = ss.getSheetByName(SHEET_VOTES);
      const sheet = ensureVotesSheet(ss);

      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === voteId) {
          sheet.deleteRow(i + 1);
          return { success: true, message: "Vote deleted", voteId: voteId };
        }
      }
      return { success: false, message: "Vote not found" };
    }

    default:
      return { success: false, message: "Unknown action" };
  }
}

/**
 * === ROUTERS ===
 */

function recordVote(contextId, token, choice) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Votes");
  const data = sheet.getDataRange().getValues();
  
  const alreadyVoted = data.some(row => row[0] === contextId && row[1] === token);
  if (alreadyVoted) {
    return { success: false, message: "Token has already voted for this context." };
  }

  sheet.appendRow([contextId, token, choice, new Date()]);
  return { success: true, message: "Vote recorded." };
}
function doGet(e) {
  try {
    const action = e.parameter.action;
    const response = handleAction(action, e.parameter);
    return sendJSON(response);
  } catch (err) {
    return sendJSON({ success: false, error: err.message });
  }
}

function doPost(e) {
  try {
    let data = {};

    // Pastikan e dan e.postData valid
    if (e && e.postData && e.postData.contents) {
      const contentType = e.postData.type || "";

      // Jika dikirim dalam format JSON
      if (contentType.indexOf("application/json") > -1) {
        try {
          data = JSON.parse(e.postData.contents);
        } catch (err) {
          data = e.parameter;
        }
      } else if (contentType.indexOf("application x-www-form-urlencoded") !== -1) {
        data = e.parameter;
      } else {
        // Jika dikirim sebagai form-data / urlencoded
        data = e.parameter;
      }
    } else {
      data = e ? e.parameter : {};
    }

    // Pastikan action ada
    const action = data.action || "undefined";
    const response = handleAction(action, data);

    return sendJSON(response);

  } catch (err) {
    Logger.log("ERROR: " + err.message);
    return sendJSON({ success: false, error: err.message });
  }
}
function isValidToken(token) {
  if (!token) return false;
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tokens");
    if (!sheet) {
      Logger.log("Tokens sheet not found");
      return false;
    }

    if (sheet.getLastRow() < 2) return false;
    const headerMap = getTokenHeaderMap(sheet);
    const tokenIndex = headerMap.token === undefined ? 0 : headerMap.token;
    const data = sheet.getRange(2, tokenIndex + 1, sheet.getLastRow() - 1, 1).getValues();
    const tokenList = data.flat().map(value => String(value).trim()).filter(Boolean);

    const valid = tokenList.includes(token.trim());
    Logger.log("Token " + token + " valid? " + valid);
    return valid;
  } catch (err) {
    Logger.log("Error reading Tokens sheet: " + err.message);
    return false;
  }
}

function getTokenHeaderMap(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return {};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headers.forEach((header, index) => {
    if (header) headerMap[String(header).trim().toLowerCase()] = index;
  });
  return headerMap;
}