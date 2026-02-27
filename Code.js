function doGet() {
  return HtmlService.createTemplateFromFile('Login').evaluate()
    .setTitle("EKR LLD Chatting App")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// === Load other pages ===
function loadPage(pageName) {
  const template = HtmlService.createTemplateFromFile(pageName);
  return template.evaluate().getContent();
}

// === CONSTANTS ===
const SHEET_ID = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";
const REG_SHEET = "xxxxxxxxxxxx";
const CHATID_SHEET = "xxxxxxx";
const P_CHAT_SHEET = "xxxxx";

// Temp storage for OTPs
var otpStore = {}

// ============ VALIDATIONS ============

// Email format check
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// Check if email already exists
function checkEmail(email) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(REG_SHEET);
  const data = sheet.getRange(2, 2, sheet.getLastRow(), 1).getValues().flat();
  return data.includes(email);
}

// Check username availability
function checkUsername(username) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(REG_SHEET);
  const data = sheet.getRange(2, 3, sheet.getLastRow(), 1).getValues().flat();
  return !data.includes(username);   // TRUE = available
}


function sendOtp(email) {
  if (!isValidEmail(email)) return { success: false, msg: "Invalid Email Format" };
  if (checkEmail(email)) return { success: false, msg: "Email Already Registered" };

  const otp = Math.floor(100000 + Math.random() * 900000).toString();

  // Store OTP for 10 minutes in cache
  const cache = CacheService.getUserCache();
  const subject = "Your OTP for EKR LLD Chatting App";
  const body = "Hello,\n\n" +
               "You have started registration in EKR LLD Chatting App.\n\n" +
               "This is your One Time Password (OTP): " + otp + "\n\n" +
               "Enter the OTP to proceed further.\n\n" +
               "Thanks and Regards,\n" +
               "EKR LLD Chatting App";
  cache.put(email, otp, 600); // 600 sec = 10 min

  // Send OTP mail
  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body,
    name: "EKR LLD Chatting App"
  });
  return { success: true, msg: "OTP Sent Successfully" };
}


function verifyOtp(email, otp) {
  const cache = CacheService.getUserCache();
  const storedOtp = cache.get(email);

  if (storedOtp && storedOtp === otp) {
    // OTP matched → clear it so it can't be reused
    cache.remove(email);
    return { success: true, msg: "User Authenticated" };
  }
  return { success: false, msg: "Invalid or Expired OTP" };
}

function createAccount(name, email, username, password) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const regSheet = ss.getSheetByName(REG_SHEET);

  if (!checkUsername(username)) {
    return { success: false, msg: "Username already exists" };
  }

  // Append user entry in Registration sheet
  regSheet.appendRow([name, email, username, password]);

  // --- Update PChats header ---
  const pchat = ss.getSheetByName(P_CHAT_SHEET);
  if (!pchat) {
    // Create PChats sheet with first header row
    const s = ss.insertSheet(P_CHAT_SHEET);
    s.appendRow([username, "Timestamp", "Status"]);
  } else {
    const lastCol = pchat.getLastColumn();
    pchat.getRange(1, lastCol + 1).setValue(username);
    pchat.getRange(1, lastCol + 2).setValue("Timestamp");
    pchat.getRange(1, lastCol + 3).setValue("Status");
  }

  return { success: true, msg: "Account Created Successfully" };
}



// Send OTP for Forgot Password
function forgotPassword(email) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(REG_SHEET);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues(); // Name, Email, Username, Password

  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === email) {
      const username = data[i][2]; // Fetch username from sheet
      const otp = Math.floor(100000 + Math.random() * 900000); // 6-digit OTP

      // Store OTP in cache for 10 minutes
      const cache = CacheService.getUserCache();
      cache.put(email, otp.toString(), 600); // 600 sec = 10 min

      // Compose personalized email
      const subject = "Password Reset OTP for EKR LLD Chatting App";
      const body = "Dear " + username + ",\n\n" +
                   "You have requested to reset your password for EKR LLD Chatting App.\n\n" +
                   "This is your One Time Password (OTP): " + otp + "\n\n" +
                   "Enter the OTP to proceed with creating a new password.\n\n" +
                   "If you did not request this, please ignore this email.\n\n" +
                   "Thanks and Regards,\n" +
                   "EKR LLD Chatting App";

      // Send OTP mail
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body,
        name: "EKR LLD Chatting App"
      });

      return { success: true, msg: "OTP sent to your email" };
    }
  }

  return { success: false, msg: "Account does not exist" };
}

// Verify OTP for Forgot Password
function ForgotPassVerOTP(email, otpInput) {
  const cache = CacheService.getUserCache();
  const storedOtp = cache.get(email);

  if (storedOtp && storedOtp === otpInput) {
    cache.remove(email); // OTP matched → remove it
    return { success: true, msg: "OTP verified" };
  }
  return { success: false, msg: "Invalid or Expired OTP" };
}

// Change Password
function changePassword(email, newPass) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(REG_SHEET);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues(); // Name, Email, Username, Password

  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === email) { // Match by email
      sheet.getRange(i + 2, 4).setValue(newPass); // Update password column
      return { success: true, msg: "Password changed successfully" };
    }
  }
  return { success: false, msg: "Account not found" };
}

/**
 * Verifies user login with username or email.
 * Returns { valid: true, username: "verifiedUsername" } on success.
 */
function loginUser(usernameOrEmail, password) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(REG_SHEET);
  if (!sheet) return { valid: false };

  const data = sheet.getDataRange().getValues(); // name | email | username | password

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[0];
    const email = String(row[1] || "").trim();
    const userName = String(row[2] || "").trim();
    const pass = String(row[3] || "").trim();

    // match username or email (case-insensitive)
    if (usernameOrEmail.toLowerCase() === email.toLowerCase() ||
        usernameOrEmail.toLowerCase() === userName.toLowerCase()) {
      
      if (password === pass) {
        // return verified username to frontend
        return {
          valid: true,
          username: userName
        };
      } else {
        return { valid: false };
      }
    }
  }

  return { valid: false };
}

/**
 * Logs each successful login to "Logins" sheet with:
 * username, deviceId, ip, timestamp
 */
function logLogin(username, deviceId, ip, timestamp) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("Logins");

  if (!sheet) {
    sheet = ss.insertSheet("Logins");
    sheet.appendRow(["Username", "DeviceID", "IP", "Timestamp"]);
  }

  sheet.appendRow([username, deviceId, ip, timestamp]);
  return true;
}




function getUserData() {
  const userProps = PropertiesService.getUserProperties();
  return {
    username: userProps.getProperty("LOGGED_USER") || "",
  };
}


// --- search users by username substring (case-insensitive) ---
function searchUsers(query) {
  query = (query || "").toLowerCase().trim();
  if (query.length < 3) return [];

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(REG_SHEET);
  const last = sheet.getLastRow();
  if (last < 2) return [];

  // Name | Email | Username in columns A,B,C
  const rows = sheet.getRange(2,1,last-1,3).getValues();
  const results = [];
  for (let i=0;i<rows.length;i++){
    const name = (rows[i][0] || "").toString();
    const username = (rows[i][2] || "").toString();
    if (!username) continue;
    if (username.toLowerCase().indexOf(query) !== -1) {
      results.push({ name: name, username: username });
      if (results.length >= 50) break; // limit
    }
  }
  return results;
}

// --- generate random 8-char alphanumeric ---
function generateChatId() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let s = "";
  for (let i=0;i<8;i++) s += chars.charAt(Math.floor(Math.random()*chars.length));
  return s;
}

// --- check if chat exists between two users (any order) ---
function findExistingChat(u1, u2) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(CHATID_SHEET);
  if (!sheet) return null;

  const last = sheet.getLastRow();
  if (last < 2) return null;

  const rows = sheet.getRange(2, 1, last - 1, 3).getValues();

  let foundChatId = null;

  for (let i = 0; i < rows.length; i++) {
    const a = String(rows[i][0] || "");
    const b = String(rows[i][1] || "");
    const id = String(rows[i][2] || "");

    // always normalize pair
    const p1 = [a, b].sort().join("-");
    const p2 = [u1, u2].sort().join("-");

    if (p1 === p2) {
      foundChatId = id;
      break;
    }
  }

  return foundChatId;
}


// --- create a sheet named chatId and append initial row ---
function createChatSheet(chatId, starter) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  // create only if not exists
  if (ss.getSheetByName(chatId)) return;
  const newSh = ss.insertSheet(chatId);
  // header optional: Sender | Message | Timestamp
  newSh.appendRow(["Sender","Message","Timestamp"]);
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
  newSh.appendRow([starter, "Started Chat", ts]);
}

// --- append to PChats for both users ---
function updatePChat(user1, user2) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(P_CHAT_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(P_CHAT_SHEET);
    sheet.appendRow([""]); // initialize row1 empty
  }

  const ts = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "MM/dd/yyyy HH:mm:ss"
  );

  // ------------------------------------------
  // helper → ensures user column exists
  // returns { userCol: number, tsCol: number }
  // ------------------------------------------
  function ensureUserColumn(username) {
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let colIndex = header.indexOf(username) + 1;

    if (colIndex === 0) {
      // create new username + timestamp columns at end
      colIndex = header.length + 1;
      sheet.getRange(1, colIndex).setValue(username);
      sheet.getRange(1, colIndex + 1).setValue("TS_" + username);
    }

    return { userCol: colIndex, tsCol: colIndex + 1 };
  }

  // create/find both user columns
  const u1 = ensureUserColumn(user1);
  const u2 = ensureUserColumn(user2);

  // ------------------------------------------
  // update entries for (owner → chattingWith)
  // ------------------------------------------
  function updateEntry(owner, other, ownerCols) {
    const { userCol, tsCol } = ownerCols;
    const lastRow = sheet.getLastRow();
    const colValues = sheet.getRange(2, userCol, lastRow - 1).getValues().flat();

    let row = colValues.indexOf(other);
    if (row >= 0) {
      // Row exists → update timestamp (row+2 because sheet starts at row2)
      sheet.getRange(row + 2, tsCol).setValue(ts);
    } else {
      // Append new entry
      sheet.getRange(lastRow + 1, userCol).setValue(other);
      sheet.getRange(lastRow + 1, tsCol).setValue(ts);
    }
  }

  // update both directions
  updateEntry(user1, user2, u1);
  updateEntry(user2, user1, u2);

  return { success: true };
}

/**
 * Create or open a chat between loggedUser and chatWith.
 * Uses existing createChatSheet and updatePChat functions.
 */
function newChat(loggedUser, chatWith) {
  if (!loggedUser || !chatWith) {
    return { success: false, msg: "Invalid users" };
  }
  if (loggedUser === chatWith) {
    return { success: false, msg: "Cannot start chat with yourself" };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);

  // --- ensure ChatID sheet exists ---
  let chatIdSheet = ss.getSheetByName(CHATID_SHEET);
  if (!chatIdSheet) {
    chatIdSheet = ss.insertSheet(CHATID_SHEET);
    chatIdSheet.appendRow(["UserA", "UserB", "ChatID"]);
  }

  // --- check if chat already exists ---
  const data = chatIdSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const u1 = String(data[i][0] || "");
    const u2 = String(data[i][1] || "");
    const cid = String(data[i][2] || "");

    if ([u1, u2].sort().join("-") === [loggedUser, chatWith].sort().join("-")) {
      // chat exists → return existing chatId
      return { success: true, chatId: cid, msg: "Chat exists" };
    }
  }

  // --- generate new 8-char alphanumeric chatId ---
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let newChatId = "";
  do {
    newChatId = "";
    for (let i = 0; i < 8; i++) {
      newChatId += chars.charAt(Math.floor(Math.random() * chars.length));
    }
  } while (ss.getSheetByName(newChatId)); // ensure unique

  // --- append new chat to ChatID sheet ---
  chatIdSheet.appendRow([loggedUser, chatWith, newChatId]);

  // --- create new chat sheet using existing function ---
  createChatSheet(newChatId, loggedUser);

  // --- update PChat entries using existing function ---
  updatePChat(loggedUser, chatWith);

  return { success: true, chatId: newChatId, msg: "New chat created" };
}


// --- main: startChat(targetUsername) ---
function startChat(loggedUser, targetUsername, chatIdFromClient) {
  // Validate inputs (replaces reading from UserProperties)
  const logged = loggedUser;
  if (!logged) return { success:false, msg:"Not signed in" };
  if (!targetUsername) return { success:false, msg:"Invalid user" };
  if (targetUsername === logged) return { success:false, msg:"Cannot start chat with yourself" };

  const u1 = logged.toString();
  const u2 = targetUsername.toString();

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let chatIdsSheet = ss.getSheetByName(CHATID_SHEET);
  if (!chatIdsSheet) {
    chatIdsSheet = ss.insertSheet(CHATID_SHEET);
    chatIdsSheet.appendRow(["UserA","UserB","ChatID"]);
  }

  // ---- store context even for existing chat (no userProps usage)
  const existing = findExistingChat(u1, u2);
  if (existing) {
    // NOTE: previously userProps were set here. Frontend now stores CHAT_WITH/CHAT_ID in localStorage.
    return { success:true, msg:"Chat exists", chatId: existing };
  }

  // new chat creation
  let chatId = chatIdFromClient ? String(chatIdFromClient) : null;

  // If provided chatId collides with an existing sheet, ignore it
  if (chatId) {
    try {
      if (ss.getSheetByName(chatId)) {
        chatId = null;
      }
    } catch (e) {
      chatId = null;
    }
  }

  let attempts = 0;
  do {
    if (!chatId) chatId = generateChatId();
    attempts++;
    if (attempts > 50) break;
  } while (ss.getSheetByName(chatId));

  chatIdsSheet.appendRow([u1, u2, chatId]);
  createChatSheet(chatId, u1);
  updatePChat(u1, u2);

  // NOTE: previously userProps were set here for CHAT_WITH / CHAT_ID.
  // Frontend must now write these values to localStorage after receiving the response.

  return { success:true, msg:"Chat created", chatId: chatId };
}



function getChatId(loggedUser, chatWith) {
  if (!loggedUser || !chatWith) return null;

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(CHATID_SHEET);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  // headers: UserA | UserB | ChatID
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const u1 = row[0];
    const u2 = row[1];
    const cid = row[2];

    if (
      (u1 === loggedUser && u2 === chatWith) ||
      (u1 === chatWith && u2 === loggedUser)
    ) {
      return cid;   // existing chat found
    }
  }

  return null; // no matching chat
}


/**
 * Reliably retrieves the email of the active user.
 * NOTE: This requires the deployment to be set to "Execute as: User accessing the web app" 
 * and requires the appropriate OAuth scopes.
 */
function getActiveUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    // Return null if session context is unavailable (e.g., app not deployed correctly)
    Logger.log("Error getting active user email: " + e);
    return null;
  }
}


// =======================
// FETCH CHAT MESSAGES
// =======================
function getChatMessages(chatId) {
  try {
    if (!chatId) return [];

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const chatSheet = ss.getSheetByName(chatId);
    if (!chatSheet) {
      Logger.log('getChatMessages: sheet not found: ' + chatId);
      return [];
    }

    const lastRow = chatSheet.getLastRow();
    if (lastRow < 1) {
      return [];
    }

    // read only used rows and first 4 columns (sender, message, timestamp, status)
    const data = chatSheet.getRange(1, 1, lastRow, 4).getValues();

    // Filter out rows with empty sender AND empty message (skip blank rows)
    const messages = data
      .map((r, index) => {
        const sender = r[0] ? String(r[0]).trim() : "";
        const message = r[1] != null ? String(r[1]) : "";
        const rawTs = r[2];
        const status = r[3] ? String(r[3]).trim() : "";

        // normalize timestamp to string
        let tsStr = "";
        if (rawTs instanceof Date) {
          tsStr = Utilities.formatDate(rawTs, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        } else if (rawTs) {
          tsStr = String(rawTs);
        }

        return {
          sender: sender,
          message: message,
          timestamp: tsStr,
          status: status,
          __rowIndex: index + 1 // helpful if you need debugging or updates
        };
      })
      .filter(m => (m.sender && m.sender.length) || (m.message && m.message.trim().length)); // at least one of these must be non-empty

    // optional: ensure stable order (sheet order is top→bottom)
    return messages;
  } catch (err) {
    Logger.log("getChatMessages error: " + err.toString());
    return [];
  }
}


function debugChatRaw(chatId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const s = ss.getSheetByName(chatId);
  if (!s) return { ok:false, msg: "no sheet" };
  const lastRow = s.getLastRow();
  const raw = s.getRange(1,1,lastRow,4).getValues();
  return { ok:true, lastRow: lastRow, raw: raw };
}



// =======================
// APPEND MESSAGE
// =======================


function appendMessage(chatId, sender, receiver, message, status) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tsStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");

  // 1) Append chat message
  const chatSheet = ss.getSheetByName(chatId);
  if (!chatSheet) throw new Error("Chat sheet not found: " + chatId);

  chatSheet.appendRow([sender, message, tsStr, status || "unseen"]);

  // 2) Determine receiver if not passed
  if (!receiver) {
    const p = chatId.split("-");
    receiver = (p[0] === sender) ? p[1] : p[0];
  }

  // 3) Update P_CHAT status for receiver's column
  const pchat = ss.getSheetByName(P_CHAT_SHEET);
  if (!pchat) return;

  const lastCol = pchat.getLastColumn();
  const headers = pchat.getRange(1, 1, 1, lastCol).getValues()[0];

  // Locate receiver column in header
  let receiverCol = -1;
  for (let c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === receiver) {
      receiverCol = c + 1; // convert 0-based to sheet column index
      break;
    }
  }
  if (receiverCol === -1) return; // receiver not present → do nothing

  // Receiver block = 3 columns: Name | Timestamp | Status
  const nameCol = receiverCol;
  const tsCol = receiverCol + 1;
  const statusCol = receiverCol + 2;

  const lastRow = pchat.getLastRow();
  if (lastRow < 2) return;

  // Search sender row in receiver's block (Name column)
  const names = pchat.getRange(2, nameCol, lastRow - 1, 1).getValues();
  let senderRow = -1;
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === sender) {
      senderRow = i + 2; // convert to sheet row
      break;
    }
  }

  if (senderRow !== -1) {
    pchat.getRange(senderRow, tsCol).setValue(tsStr);
    pchat.getRange(senderRow, statusCol).setValue("Unseen");
  }

  // 4) New feature: Update sender's column for receiver (timestamp only)
  let senderCol = -1;
  for (let c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === sender) {
      senderCol = c + 1;
      break;
    }
  }
  if (senderCol === -1) return; // sender column not found

  const receiverNames = pchat.getRange(2, senderCol, lastRow - 1, 1).getValues();
  let receiverRow = -1;
  for (let i = 0; i < receiverNames.length; i++) {
    if (String(receiverNames[i][0]).trim() === receiver) {
      receiverRow = i + 2;
      break;
    }
  }

  if (receiverRow !== -1) {
    pchat.getRange(receiverRow, senderCol + 1).setValue(tsStr); // only update timestamp
    // status is left untouched
  }
}


// =======================
// MARK LAST MESSAGE SEEN
// =======================
function markLastMessageSeen(chatId, chatWith, loggedInUser) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // -----------------------
  // 1. Update in chat sheet
  // -----------------------
  const chatSheet = ss.getSheetByName(chatId);
  if (chatSheet) {
    const data = chatSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 0; i--) { 
      if (data[i][0] === chatWith && data[i][3] !== "Seen") { 
        chatSheet.getRange(i + 1, 4).setValue("Seen"); 
        break;
      }
    }
  }

  // -----------------------
  // 2) Update in P_CHAT_SHEET (same pattern as appendMessage)
  const pchat = ss.getSheetByName(P_CHAT_SHEET);
  if (!pchat) return;

  const lastCol = pchat.getLastColumn();
  const headers = pchat.getRange(1, 1, 1, lastCol).getValues()[0];

  // find loggedInUser's column
  let userCol = -1;
  for (let c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === loggedInUser) {
      userCol = c + 1; // convert index
      break;
    }
  }
  if (userCol === -1) return;

  // loggedInUser block: 3 columns
  const nameCol = userCol;
  const tsCol = userCol + 1;
  const statusCol = userCol + 2;

  const lastRow = pchat.getLastRow();
  if (lastRow < 2) return;

  // find chatWith under loggedInUser block
  const names = pchat.getRange(2, nameCol, lastRow - 1, 1).getValues();

  let targetRow = -1;
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === chatWith) {
      targetRow = i + 2;
      break;
    }
  }

  if (targetRow === -1) return;

  // update ONLY status → "Seen"
  pchat.getRange(targetRow, statusCol).setValue("Seen");
}

// =======================
// UPDATE USER TIMESTAMP
// =======================
function updateUserTimestamp(username) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Registration");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === username) {    // column C = UserName (0=A,1=B,2=C)
      sheet.getRange(i + 1, 5).setValue(new Date()); // column E = Last Seen
      return true;
    }
  }

  return false; // user not found
}

// =======================
// CHECK USER ONLINE STATUS
// =======================
function getUserStatus(username) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(REG_SHEET);
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 0; i < data.length; i++) {
    if (data[i][2] === username) {
      const tsStr = data[i][4];
      if (!tsStr) return false;
      const ts = new Date(tsStr);
      return (now - ts) / 1000 <= 30;
    }
  }
  return false;
}

// =======================
// REQUIRED FOR CHATSCREEN:
// GET ONLINE + LAST SEEN
// =======================
function getOnlineInfo(username) {
  const online = getUserStatus(username);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(REG_SHEET);
  const data = sheet.getDataRange().getValues();

  let lastSeen = "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === username) {
      lastSeen = data[i][4] || "";
      break;
    }
  }

  // convert to ISO string if valid date
  if (lastSeen instanceof Date && !isNaN(lastSeen)) {
    lastSeen = lastSeen.toISOString();
  }

  return {
    online,
    lastSeen
  };
}

function getPChatsForUser(logged) {
  
  if (!logged) return [];

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const pchat = ss.getSheetByName(P_CHAT_SHEET);
  if (!pchat) return [];

  const lastCol = pchat.getLastColumn();
  const lastRow = pchat.getLastRow();
  if (lastCol < 1 || lastRow < 2) return [];

  // read header row
  const headers = pchat.getRange(1, 1, 1, lastCol).getValues()[0];

  // find logged user column index
  let userCol = -1;
  for (let c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === logged) {
      userCol = c + 1;   // convert to sheet column
      break;
    }
  }
  if (userCol === -1) return [];

  // user block = 3 columns
  const nameCol = userCol;
  const tsCol   = userCol + 1;
  const stCol   = userCol + 2;

  // read user block rows (names, timestamps, statuses)
  const rowCount = lastRow - 1;
  const names = pchat.getRange(2, nameCol, rowCount, 1).getValues();
  const timestamps = pchat.getRange(2, tsCol, rowCount, 1).getValues();
  const statuses = pchat.getRange(2, stCol, rowCount, 1).getValues();

  const result = [];

  for (let i = 0; i < rowCount; i++) {
    const uname = String(names[i][0] || "").trim();
    if (!uname) continue; // skip blanks

    const ts = String(timestamps[i][0] || "").trim();
    const st = String(statuses[i][0] || "").trim();

    // get full name from registration sheet
    const fullName = lookupName(uname);

    result.push({
      name: fullName || uname,  // fallback to username
      username: uname,
      timestamp: ts,
      status: st
    });
  }

  return result;
}

function lookupName(username) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const reg = ss.getSheetByName(REG_SHEET);
  if (!reg) return "";

  const data = reg.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === username) {  // username column
      return data[i][0];  // name column
    }
  }
  return "";
}


function getMyProfileData(loggedUser) {
  if (!loggedUser) {
    return {
      username: "",
      name: "",
      email: "",
      totalChats: 0
    };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ===== REGISTRATION SHEET =====
  const reg = ss.getSheetByName(REG_SHEET);
  let name = "";
  let email = "";

  if (reg) {
    const data = reg.getDataRange().getValues(); // name | email | username
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[2] || "").trim() === loggedUser) { // username in col 3
        name = row[0] || "";
        email = row[1] || "";
        break;
      }
    }
  }

  // ===== TOTAL CHATS FROM P_CHAT_SHEET =====
  let totalChats = 0;
  const pchat = ss.getSheetByName(P_CHAT_SHEET);

  if (pchat) {
    const lastCol = pchat.getLastColumn();
    const lastRow = pchat.getLastRow();

    if (lastCol >= 1 && lastRow >= 2) {
      const headers = pchat.getRange(1, 1, 1, lastCol).getValues()[0];

      // Find column for loggedUser
      const userColIndex = headers.findIndex(h => String(h || "").trim() === loggedUser);
      if (userColIndex !== -1) {
        const names = pchat.getRange(2, userColIndex + 1, lastRow - 1, 1).getValues();
        totalChats = names.filter(r => r[0] && String(r[0]).trim() !== "").length;
      }
    }
  }

  return {
    username: loggedUser,
    name: name,
    email: email,
    totalChats: totalChats
  };
}
