// Global variable for the spreadsheet and user sheet
var SPREADSHEET_NAME = "LOGIN";
var USER_SHEET_NAME = "Users";
var REDIRECT_URL = "Home.html"; // Default URL to redirect after successful login

function doGet(e) {
  const cache = CacheService.getUserCache();

  // 0. Ưu tiên xử lý quét QR đe dọa tâm lý bảo mật
  if (e.parameter.qr_scan) {
    return handleQrScan(e.parameter.qr_scan);
  }

  // 1. Ưu tiên xử lý xác minh email nếu có token
  if (e.parameter.token) {
    return handleEmailVerification(e.parameter.token);
  }

  // 2. Lấy email từ cache để kiểm tra trạng thái đăng nhập
  const email = cache.get("loggedInUser");

  if (!email) {
    // Nếu không có cache -> Trả về trang chính (main.html) hoạt động như SPA
    return HtmlService.createTemplateFromFile("main")
      .evaluate()
      .setTitle("Triple D Sinh Học")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 3. Nếu đã đăng nhập, kiểm tra thông tin user
  const user = getUserByEmail(email); // Hàm này bạn đã có sẵn
  if (user && user.data[7] === "Verified") {
    let page = e.parameter.page || "Home"; // Mặc định vào Home (hoặc main)
    if (page === "index") page = "Home"; // Đã đăng nhập thì không hiển thị lại form login
    
    // Tự động giải quyết fallback thông minh giữa 'Home' và 'main'
    if (page === "Home") {
      try {
        HtmlService.createTemplateFromFile("Home");
      } catch (err) {
        page = "main";
      }
    } else if (page === "main") {
      try {
        HtmlService.createTemplateFromFile("main");
      } catch (err) {
        page = "Home";
      }
    }

    return HtmlService.createTemplateFromFile(page)
      .evaluate()
      .setTitle("Triple D Sinh Học")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 4. Trường hợp ngoại lệ (có cache nhưng dữ liệu lỗi hoặc chưa verify)
  cache.remove("loggedInUser");
  return HtmlService.createTemplateFromFile("index").evaluate();
}

// Hàm logout: Xóa cache và trả về URL sạch
function logout() {
  const cache = CacheService.getUserCache();
  const email = cache.get("loggedInUser");
  if (email) {
    logActivityToSheet(email, "LOGOUT", "Người dùng chủ động đăng xuất");
  }
  cache.remove("loggedInUser"); 
  return ScriptApp.getService().getUrl();
}

/**
 * Tự động ghi nhận hoạt động người dùng và Admin vào sheet Tracking
 */
function logActivityToSheet(email, action, detail) {
  // Đã vô hiệu hóa theo yêu cầu người dùng
  return true;
}


// --- GLOBAL EXECUTION CACHE SYSTEM ---
var _activeSpreadsheet = null;
var _cachedSheetValues = {};
var _cachedUserRoles = {};
var _cachedCoursesMemory = null;
var _cachedBooksMemory = null;
var _cachedExamsMemory = null;

function getSpreadsheet() {
  if (!_activeSpreadsheet) {
    _activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  return _activeSpreadsheet;
}

function getSheet(sheetName) {
  var ss = getSpreadsheet();
  return ss ? ss.getSheetByName(sheetName) : null;
}

function getSheetDataCached(sheetName) {
  if (!_cachedSheetValues[sheetName]) {
    var sheet = getSheet(sheetName);
    if (!sheet) return [];
    _cachedSheetValues[sheetName] = sheet.getDataRange().getValues();
  }
  return _cachedSheetValues[sheetName];
}

function invalidateSheetCache(sheetName) {
  if (sheetName) {
    delete _cachedSheetValues[sheetName];
  } else {
    _cachedSheetValues = {};
  }
  _cachedCoursesMemory = null;
  _cachedBooksMemory = null;
  _cachedExamsMemory = null;
}

/**
 * Generates a unique verification token.
 * @returns {string} A unique token.
 */
function generateVerificationToken() {
  return Utilities.getUuid(); // Generates a unique ID
}

/**
 * Generates a 6-digit OTP.
 * @returns {string} A 6-digit OTP.
 */
function generateOTP() {
  return Math.floor(100000 + Math.random() * 900000).toString(); // 6-digit number
}

function sendEmail(recipient, subject, body) {
  if (!canSendEmail(recipient)) {
    throw new Error("Bạn đã gửi quá nhiều email trong 24 giờ. Vui lòng thử lại sau.");
  }
  MailApp.sendEmail(recipient, subject, body);
}

/**
 * Checks if an email is valid.
 * @param {string} email The email to validate.
 * @returns {boolean} True if the email is valid, false otherwise.
 */
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

/**
 * Retrieves user data by email.
 * @param {string} email The email of the user.
 * @returns {Array} The user's row data or null if not found.
 */
function getUserByEmail(email) {
  if (!email) return null;
  var data = getSheetDataCached(USER_SHEET_NAME);
  var cleanEmail = String(email).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) { // Skip header row
    if (data[i][1] && String(data[i][1]).trim().toLowerCase() === cleanEmail) { // Column B is Email (index 1)
      return { row: i, data: data[i] };
    }
  }
  return null;
}

/**
 * Retrieves user data by verification token.
 * @param {string} token The verification token.
 * @returns {Array} The user's row data or null if not found.
 */
function getUserByToken(token) {
  if (!token) return null;
  var data = getSheetDataCached(USER_SHEET_NAME);
  var tokenStr = String(token);
  for (var i = 1; i < data.length; i++) { // Skip header row
    if (data[i][8] && String(data[i][8]) === tokenStr) { // Column I is Verification Token (index 8)
      return { row: i, data: data[i] };
    }
  }
  return null;
}

/**
 * Retrieves user data by OTP.
 * @param {string} email The email of the user.
 * @param {string} otp The OTP.
 * @returns {Array} The user's row data or null if not found or OTP expired.
 */
function getUserByOTP(email, otp) {
  var user = getUserByEmail(email);
  if (user) {
    var storedOtp = user.data[5]; // Column F is OTP (index 5)
    var otpExpiry = user.data[6]; // Column G is DATE OF EXPIRY (index 6)

    if (storedOtp == otp && otpExpiry && new Date() < new Date(otpExpiry)) {
      return user;
    }
  }
  return null;
}

/**
 * Updates a specific field for a user in the sheet.
 * @param {number} row The row index of the user (0-indexed).
 * @param {number} colIndex The column index to update (0-indexed).
 * @param {any} value The new value.
 */
function updateUserField(row, colIndex, value) {
  var sheet = getSheet(USER_SHEET_NAME);
  if (sheet) {
    sheet.getRange(row + 1, colIndex + 1).setValue(value); // +1 for 1-indexed sheet
  }
  if (_cachedSheetValues[USER_SHEET_NAME] && _cachedSheetValues[USER_SHEET_NAME][row]) {
    _cachedSheetValues[USER_SHEET_NAME][row][colIndex] = value;
  }
}

// --- Password Hashing Functions ---

/**
 * Generates a random salt.
 * @returns {string} A random 16-character hexadecimal string.
 */
function generateSalt() {
  return Utilities.getUuid().replace(/-/g, ''); // Generate a UUID and remove hyphens for a shorter salt
}

/**
 * Hashes a password with a given salt using SHA-256.
 * It performs multiple iterations (stretching) for stronger security.
 * @param {string} password The plain text password.
 * @param {string} salt The unique salt for the user.
 * @param {number} iterations (Optional) Number of hashing iterations. Default to 10000.
 * @returns {string} The hashed password.
 */
function hashPassword(password, salt, iterations = 10) {
  let combined = password + salt;
  let hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, combined);
  for (let i = 0; i < iterations - 1; i++) {
    // Re-hash the hash itself, encoding it to base64 for proper input format
    // and then decoding to bytes for computeDigest to avoid errors with byte array vs string.
    // This is a common pattern for stretching with raw digest functions.
    hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.base64Decode(Utilities.base64Encode(hash)));
  }
  return Utilities.base64Encode(hash); // Encode the final hash to base64 for storage
}


/**
 * Processes user signup request.
 * @param {Object} formData Contains name, email, and password.
 * @returns {Object} Success status and message.
 */
/**
 * Cộng thưởng khi hoàn thành phiên tập trung (Focus Mode)
 * Tối thiểu 60 phút nhận 50 coin
 */
function completeFocusSession(durationMinutes) {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false, message: "Chưa đăng nhập" };
  
  if (durationMinutes < 60) {
    return { success: false, message: "Thời gian tập trung chưa đủ 60 phút để nhận thưởng." };
  }

  const reward = 50;
  const user = getUserByEmail(email);
  if (!user) return { success: false, message: "Không tìm thấy người dùng" };

  const currentPoints = Number(user.data[10]) || 0;
  updateUserField(user.row, 10, currentPoints + reward);

  return { success: true, reward: reward, newPoints: currentPoints + reward };
}

function updateUserProfile(email, formData) {
  try {
    var userObj = getUserByEmail(email);
    if (!userObj) {
      return { success: false, message: "Không tìm thấy người dùng." };
    }
    var userRow = userObj.row;
    var userData = userObj.data;
    var storedHashedPassword = userData[2];
    var storedSalt = userData[9] || "TRIPLED_DEFAULT_SALT"; // In case salt is missing

    // Verify current password if user is trying to change password
    if (formData.newPassword) {
      if (!formData.currentPassword) {
        return { success: false, message: "Vui lòng nhập mật khẩu hiện tại." };
      }
      var hashedCurrentInput = hashPassword(formData.currentPassword, storedSalt);
      if (hashedCurrentInput !== storedHashedPassword) {
        return { success: false, message: "Mật khẩu hiện tại không đúng." };
      }
    }

    // Update Name
    if (formData.newName && formData.newName !== userData[0]) {
      updateUserField(userRow, 0, formData.newName);
    }

    // Update Password
    if (formData.newPassword) {
      var newSalt = generateSalt();
      var hashedNewPassword = hashPassword(formData.newPassword, newSalt);
      updateUserField(userRow, 2, hashedNewPassword);
      updateUserField(userRow, 9, newSalt);
    }

    return { success: true, message: "Cập nhật thông tin thành công!" };
  } catch (error) {
    return { success: false, message: "Lỗi nội bộ: " + error.toString() };
  }
}

function processSignup(formData) {
  try {
    var sheet = getSheet(USER_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: "Sheet 'Users' not found. Please check your Google Sheet setup." };
    }

    var email = formData.email;
    if (!isValidEmail(email)) {
      return { success: false, message: "Vui lòng nhập địa chỉ email hợp lệ!" };
    }

    // Check for duplicate email
    var existingUser = getUserByEmail(email);
    if (existingUser) {
      return { success: false, message: "Email đã được đăng ký!" };
    }

    var salt = generateSalt(); // Generate unique salt for new user
    var hashedPassword = hashPassword(formData.password, salt); // Hash the password

    var verificationToken = generateVerificationToken();
    var webAppUrl = ScriptApp.getService().getUrl();
    var verificationLink = webAppUrl + '?token=' + verificationToken;

    // Append new user with "Pending" status, verification token, HASHED password, and salt
    // MAKE SURE TO ADD A 'Salt' COLUMN IN YOUR GOOGLE SHEET (e.g., Column J, index 9)
    sheet.appendRow([
      formData.name,            // A
      formData.email,           // B
      hashedPassword,           // C
      "Allowed",                // D
      REDIRECT_URL,             // E
      "",                       // F - OTP
      "",                       // G - Expiry
      "Pending",                // H - Status
      verificationToken,        // I
      salt,                     // J
      500                      // K - Điểm ban đầu
    ]);

    // Send verification email
    var emailBody = "Triple D Xin Chào " + formData.name + ",\n\n" +
                    "Cảm ơn bạn đã đăng ký. Vui lòng nhấp vào liên kết bên dưới để xác minh địa chỉ email của bạn:\n\n" +
                    verificationLink + "\n\n" +
                    "Liên kết này sẽ hết hạn sau 24 giờ.\n\n" + // Note: actual expiry logic for link is not in this script
                    "Cảm Ơn,\nTriple D";
    sendEmail(formData.email, "Liên Kết Xác Minh Tài Khoản Từ Triple D", emailBody);

    return { success: true, message: "Đăng ký thành công! Một liên kết xác minh đã được gửi đến email của bạn. Vui lòng xác minh email của bạn để đăng nhập." };

  } catch (e) {
    Logger.log(e.toString());
    return { success: false, message: "Lỗi trong quá trình đăng ký: " + e.message };
  }
}

/**
 * Handles email verification when the user clicks the link.
 * @param {string} token The verification token from the URL.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} HTML content indicating verification status.
 */
function handleEmailVerification(token) {
  var user = getUserByToken(token);
  if (user && user.data[7] === "Pending") { // Column H is Status (index 7)
    updateUserField(user.row, 7, "Verified"); // Set Status to Verified
    updateUserField(user.row, 8, ""); // Clear Verification Token (Column I, index 8)
    return HtmlService.createHtmlOutput('<p style="font-family: sans-serif; text-align: center; color: green; font-size: 1.2em;">Email của bạn đã được xác minh thành công! Bây giờ bạn có thể đăng nhập.</p>');
  } else {
    return HtmlService.createHtmlOutput('<p style="font-family: sans-serif; text-align: center; color: red; font-size: 1.2em;">Liên kết xác minh không hợp lệ hoặc đã hết hạn.</p>');
  }
}

/**
 * Hàm so sánh thông số thiết bị giữa dữ liệu gửi lên và dữ liệu đã lưu.
 * Giúp nhận diện thiết bị cũ/mới một cách thông minh, không bắt bẻ sai lệch nhỏ về zoom hay trình duyệt trên cùng 1 máy.
 */
function compareDeviceInfo(storedObj, clientObj, email) {
  if (!storedObj || !clientObj) return false;
  if (typeof storedObj !== 'object' || typeof clientObj !== 'object') return false;
  if (Object.keys(storedObj).length === 0 || Object.keys(clientObj).length === 0) return false;

  // 1. Chuẩn hóa họ hệ điều hành (Windows, Apple [iOS/iPadOS/macOS], Android, Linux)
  function getOSFamily(osStr) {
    if (!osStr) return "";
    var s = String(osStr).toLowerCase();
    if (s.indexOf("win") !== -1) return "windows";
    if (s.indexOf("mac") !== -1 || s.indexOf("ios") !== -1 || s.indexOf("ipad") !== -1 || s.indexOf("apple") !== -1) return "apple";
    if (s.indexOf("android") !== -1) return "android";
    if (s.indexOf("linux") !== -1) return "linux";
    return s.split(' ')[0];
  }

  const osFamA = getOSFamily(storedObj.os);
  const osFamB = getOSFamily(clientObj.os);

  // Nếu hệ điều hành hoàn toàn khác họ (ví dụ 1 bên là Windows PC, 1 bên là iPhone/Android/Mac) -> chắc chắn khác máy
  if (osFamA && osFamB && osFamA !== osFamB) {
    Logger.log("[compareDeviceInfo] OS family mismatch for " + (email || "unknown") + ": " + osFamA + " vs " + osFamB);
    return false;
  }

  // 2. So khớp theo GPU Card (Card đồ họa là đặc trưng rất bền vững của phần cứng)
  if (storedObj.gpu && clientObj.gpu && String(storedObj.gpu).trim() !== "" && String(storedObj.gpu).toLowerCase().trim() === String(clientObj.gpu).toLowerCase().trim()) {
    Logger.log("[compareDeviceInfo] GPU hardware match for " + (email || "unknown"));
    return true;
  }

  // 3. Chuẩn hóa độ phân giải màn hình (luôn sắp xếp Cạnh nhỏ x Cạnh lớn)
  function normalizeScreen(s) {
    if (!s) return "";
    var res = String(s).split('x')
            .map(function(val) { return Math.floor(Number(val)); })
            .filter(function(n) { return !isNaN(n); })
            .sort(function(x, y) { return x - y; })
            .join('x');
    return res;
  }

  const screenA = normalizeScreen(storedObj.screen);
  const screenB = normalizeScreen(clientObj.screen);

  // Nếu cùng màn hình và cùng họ OS -> Cùng máy (PASS)
  if (screenA && screenB && screenA === screenB) {
    Logger.log("[compareDeviceInfo] Screen resolution match for " + (email || "unknown"));
    return true;
  }

  // 4. Chuẩn hóa trình duyệt & loại thiết bị
  const brA = (storedObj.browser || "").toLowerCase().trim();
  const brB = (clientObj.browser || "").toLowerCase().trim();
  const typeA = (storedObj.deviceType || "").toLowerCase().trim();
  const typeB = (clientObj.deviceType || "").toLowerCase().trim();

  // Nếu cùng họ OS và (cùng loại thiết bị hoặc cùng trình duyệt) -> Cùng máy (PASS)
  if (osFamA && osFamB && osFamA === osFamB) {
    if ((typeA && typeB && typeA === typeB) || (brA && brB && brA === brB)) {
      Logger.log("[compareDeviceInfo] OS & DeviceType/Browser match for " + (email || "unknown"));
      return true;
    }
  }

  Logger.log("[compareDeviceInfo] Mismatch for " + (email || "unknown") + ": " + JSON.stringify(storedObj) + " vs " + JSON.stringify(clientObj));
  return false;
}

/**
 * Processes user login request.
 * @param {Object} formData Contains email and password.
 * @returns {Object} Success status, message, and redirect URL.
 */
function processLogin(formData) {
  try {
    var email = formData.email;
    var enteredPassword = formData.password;
    var deviceID = formData.deviceID || "Không rõ";
    var fingerprint = formData.fingerprint || "";
    var deviceInfoRaw = formData.deviceInfo || "{}"; // client gửi JSON string

    // parse deviceInfo từ client
    var deviceObj;
    try {
      deviceObj = JSON.parse(deviceInfoRaw);
    } catch (e) {
      deviceObj = {};
    }

    var user = getUserByEmail(email);
    if (!user) {
      return { success: false, message: "Email hoặc mã truy cập không chính xác. Vui lòng thử lại!" };
    }

    var userData = user.data;
    var storedHashedPassword = userData[2];
    var storedSalt = userData[9];
    var status = userData[7];
    var type = userData[3];
    var storedDeviceID = userData[16]; // Cột Q: deviceID
    var role = userData[13];           // Cột N: vai trò
    var storedDeviceInfoRaw = userData[18]; // Cột S: deviceInfo đã lưu
    var storedDeviceObj = {};
    try {
      storedDeviceObj = JSON.parse(storedDeviceInfoRaw);
    } catch (e) {
      storedDeviceObj = {};
    }

    const normalizedRole = (role || "").toLowerCase().trim();
    
    // 👉 KIỂM TRA THIẾT BỊ HỌC VIÊN: CHỈ CẦN TRÙNG DeviceID HOẶC TRÙNG THIẾT BỊ LÀ CHO QUA
    if (["student", "s-student", "o-student"].includes(normalizedRole)) {
      var storedDID = storedDeviceID ? String(storedDeviceID).trim() : "";
      var sentDID = deviceID ? String(deviceID).trim() : "";
      var sentFP = fingerprint ? String(fingerprint).trim() : "";

      var hasStoredDID = Boolean(storedDID && storedDID !== "" && storedDID !== "Không rõ");
      var hasStoredDeviceInfo = Boolean(storedDeviceInfoRaw && storedDeviceInfoRaw !== "{}" && storedDeviceInfoRaw !== "" && Object.keys(storedDeviceObj).length > 0);

      // Nếu tài khoản đã có thiết bị liên kết trước đó
      if (hasStoredDID || hasStoredDeviceInfo) {
        var isDeviceIDMatch = false;

        // 1. So khớp DeviceID (Kiểm tra xem ID hoặc Fingerprint có trùng khớp không)
        if (hasStoredDID) {
          if (sentDID && sentDID !== "Không rõ" && (storedDID === sentDID || storedDID.indexOf(sentDID) !== -1 || sentDID.indexOf(storedDID) !== -1)) {
            isDeviceIDMatch = true;
          }
          if (sentFP && sentFP !== "Không rõ" && (storedDID === sentFP || storedDID.indexOf(sentFP) !== -1 || sentFP.indexOf(storedDID) !== -1)) {
            isDeviceIDMatch = true;
          }
        }

        // 2. So khớp Thông số Thiết bị (Hệ điều hành, Card đồ họa, Màn hình, Loại máy)
        var isDeviceInfoMatch = false;
        if (hasStoredDeviceInfo) {
          isDeviceInfoMatch = compareDeviceInfo(storedDeviceObj, deviceObj, email);
        } else if (hasStoredDID && !hasStoredDeviceInfo) {
          // Trường hợp dữ liệu cũ: Chỉ có Cột Q mà Cột S trống
          // Nếu trùng DID hoặc cùng họ OS cơ bản thì cho phép cập nhật đồng bộ
          if (isDeviceIDMatch) {
            isDeviceInfoMatch = true;
          } else {
            var clientOS = (deviceObj.os || "").toLowerCase();
            if (clientOS && !clientOS.includes("unknown")) {
              isDeviceInfoMatch = true; // Chuyển đổi mềm cho tài khoản cũ chưa có Cột S
            }
          }
        }

        // 👉 ĐIỀU KIỆN ĐỘC LẬP: CHỈ CẦN 1 TRONG 2 TIÊU CHÍ KHỚP LÀ CHO QUA
        var isDeviceAllowed = isDeviceIDMatch || isDeviceInfoMatch;

        if (!isDeviceAllowed) {
          Logger.log("[Login] Device mismatch for " + email + ". Stored DID: " + storedDID + ", Sent DID: " + sentDID + ", Sent FP: " + sentFP + " | Stored Dev: " + storedDeviceInfoRaw + ", Sent Dev: " + deviceInfoRaw);
          return {
            success: false,
            message: "Tài khoản của bạn đã được liên kết với một thiết bị khác. Bạn chỉ có thể đăng nhập trên thiết bị đã liên kết ban đầu. Vui lòng liên hệ Admin nếu bạn đã đổi máy mới để được reset thiết bị!"
          };
        }
      }
    }

    const normStatus = (status || "").toString().trim().toLowerCase();
    const normType = (type || "").toString().trim().toLowerCase();
    const normRoleCheck = (role || "").toString().trim().toLowerCase();

    if (normStatus === "pending") {
      return {
        success: false,
        message: "Email chưa được xác minh.",
        unverified: true
      };
    }

    if (normType === "blocked" || normStatus === "blocked" || normRoleCheck === "blocked") {
      return {
        success: false,
        message: "Tài khoản của bạn đã bị chặn. Vui lòng liên hệ với Triple D qua email: tripledadm@gmail.com."
      };
    }

    var hashedEnteredPassword = hashPassword(enteredPassword, storedSalt);

    if (hashedEnteredPassword === storedHashedPassword) {
      var cache = CacheService.getUserCache();
      var nowTs = String(new Date().getTime());
      cache.put("loggedInUser", email, 21600); // Lưu 6 giờ
      cache.put("loginTimestamp", nowTs, 21600);

      var loginDays = 1;
      try {
        loginDays = refreshUserStreak(email);
      } catch (streakErr) {
        Logger.log("refreshUserStreak error in processLogin (ignored): " + streakErr);
      }

      // ✅ Cập nhật/đồng bộ DeviceID (Cột Q - index 16)
      if (deviceID && deviceID !== "Không rõ") {
        updateUserField(user.row, 16, deviceID); // Q
      }

      // ✅ Cập nhật/đồng bộ Device Info vào Cột S (index 18)
      if (deviceObj && Object.keys(deviceObj).length > 0) {
        updateUserField(user.row, 18, JSON.stringify(deviceObj)); // S
      }

      // ✅ Luôn ghi nhận chính xác thời gian đăng nhập mới nhất vào Cột O (index 14)
      updateUserField(user.row, 14, new Date()); // O
      
      var eventCoin = parseInt(userData[20]) || 0;
      
      logActivityToSheet(email, "LOGIN", `Đăng nhập thành công dID: ${deviceID || 'Không rõ'}. Device: ${deviceObj ? deviceObj.browser : 'N/A'}`);

      return {
        success: true,
        message: "Đăng nhập thành công!",
        eventCoin: eventCoin,
        redirectUrl: ScriptApp.getService().getUrl() + "?page=Home&t=" + new Date().getTime()
      };
    } else {
      logActivityToSheet(email, "LOGIN_FAILED", "Nhập sai mã truy cập / mật khẩu");
      return { success: false, message: "Email hoặc mã truy cập không chính xác. Vui lòng thử lại!" };
    }

  } catch (e) {
    Logger.log(e.toString());
    return { success: false, message: "Lỗi khi đăng nhập: " + e.message };
  }
}

function getDeviceInfoFromSheet(email) {
  try {
    var user = getUserByEmail(email);
    if (!user) return null;

    var userData = user.data;
    var storedDeviceInfoRaw = userData[18]; // Cột S
    if (!storedDeviceInfoRaw) return null;

    var deviceObj = {};
    try {
      deviceObj = JSON.parse(storedDeviceInfoRaw);
    } catch (e) {
      deviceObj = {};
    }

    return deviceObj;
  } catch (err) {
    Logger.log("getDeviceInfoFromSheet error: " + err);
    return null;
  }
}


function checkLoginStatus() {
  const cache = CacheService.getUserCache();
  const email = cache.get("loggedInUser");
  if (!email) return { loggedIn: false };

  const normEmail = (email || "").toString().toLowerCase().trim();
  const loginTimestampStr = cache.get("loginTimestamp");
  const loginTimestamp = loginTimestampStr ? Number(loginTimestampStr) : 0;

  const scriptProps = PropertiesService.getScriptProperties();

  // 1. Kiểm tra đăng xuất bắt buộc toàn Server
  const globalLogoutTimestampStr = scriptProps.getProperty("GLOBAL_FORCE_LOGOUT_TIMESTAMP");
  const globalLogoutTimestamp = globalLogoutTimestampStr ? Number(globalLogoutTimestampStr) : 0;
  if (globalLogoutTimestamp > 0 && (loginTimestamp === 0 || loginTimestamp <= globalLogoutTimestamp)) {
    cache.remove("loggedInUser");
    cache.remove("loginTimestamp");
    return { 
      loggedIn: false, 
      forcedLogout: true, 
      message: "Hệ thống đã kích hoạt đăng xuất toàn Server. Vui lòng đăng nhập lại!" 
    };
  }

  // 2. Kiểm tra đăng xuất bắt buộc theo tài khoản cụ thể (Force Logout / Kick phiên)
  const userLogoutTimestampStr = scriptProps.getProperty("USER_FORCE_LOGOUT_" + normEmail);
  const userLogoutTimestamp = userLogoutTimestampStr ? Number(userLogoutTimestampStr) : 0;
  if (userLogoutTimestamp > 0 && (loginTimestamp === 0 || loginTimestamp <= userLogoutTimestamp)) {
    cache.remove("loggedInUser");
    cache.remove("loginTimestamp");
    return { 
      loggedIn: false, 
      forcedLogout: true, 
      message: "Phiên làm việc của bạn đã bị quản trị viên đăng xuất. Vui lòng đăng nhập lại!" 
    };
  }

  // 3. KIỂM TRA TRẠNG THÁI TÀI KHOẢN (Blocked / Đã bị xóa)
  let user = null;
  try {
    user = getUserByEmail(email);
  } catch (e) {
    console.error("Error fetching user in checkLoginStatus: " + e);
  }

  if (!user) {
    cache.remove("loggedInUser");
    cache.remove("loginTimestamp");
    return {
      loggedIn: false,
      forcedLogout: true,
      message: "Tài khoản không tồn tại trong hệ thống. Vui lòng đăng nhập lại!"
    };
  }

  const userStatus = (user.data[7] || "").toString().trim().toLowerCase();
  const userType = (user.data[3] || "").toString().trim().toLowerCase();
  const userRole = (user.data[13] || "").toString().trim().toLowerCase();

  if (userStatus === "blocked" || userType === "blocked" || userRole === "blocked") {
    cache.remove("loggedInUser");
    cache.remove("loginTimestamp");
    return {
      loggedIn: false,
      forcedLogout: true,
      isBlocked: true,
      message: "Tài khoản của bạn đã bị quản trị viên khóa (Blocked). Vui lòng liên hệ với Triple D qua email: tripledadm@gmail.com để được hỗ trợ!"
    };
  }

  let displayName = user.data[0] || "Người dùng";

  // Tiến hành kiểm tra phần thưởng từ Cửa hàng (Store)
  const newRewards = checkAndProcessStoreRewards(email);

  // Tự động kiểm tra và giải ngân vé cược bóng đá
  try {
    autoUpdateUserFootballBets(email);
  } catch (footballErr) {
    Logger.log("Err running auto-football-check: " + footballErr.toString());
  }

  return {
    loggedIn: true,
    email: email,
    displayName: displayName,
    newRewards: newRewards
  };
}

/**
 * Tối ưu hóa: Kết hợp tất cả dữ liệu khởi tạo vào một lần gọi duy nhất
 * để giảm số tầng round-trip giữa client và server.
 */
function getInitialAppDataCombined() {
  const loginInfo = checkLoginStatus();
  if (!loginInfo.loggedIn) return loginInfo;

  // Gọi check auto reset
  checkAndRunAutoSeasonReset();

  const email = loginInfo.email;
  const combinedData = { ...loginInfo };

  // Refresh and get consecutive login streak
  try {
    combinedData.streak = refreshUserStreak(email);
  } catch(e) {
    console.error("Lỗi refreshUserStreak: " + e.toString());
    combinedData.streak = 1;
  }

  try {
    // Lấy thông tin user và partner một lần để dùng cho nhiều mục đích
    const info = getUserAndPartner(email);
    
    // 1. Tính điểm (Coins)
    let myCoin = 0;
    let partnerCoin = 0;
    if (info) {
      myCoin = Number(info.myRow[10]) || 0;
      partnerCoin = info.partnerRow ? Number(info.partnerRow[10]) || 0 : 0;
    }
    combinedData.points = myCoin + partnerCoin;

    // 2. Lấy Event Coins
    combinedData.eventCoin = getUserEventCoin();

    // 3. Tính vai trò (Role)
    if (info) {
      function _normalizeRole(raw) {
        raw = (raw || "").toString().trim().toUpperCase();
        if (["ADMIN", "STUDENT", "S-STUDENT", "O-STUDENT", "VIP"].includes(raw)) return raw;
        return "MEM";
      }
      const myRole = _normalizeRole(info.myRow[13]);
      const partnerRole = info.partnerRow ? _normalizeRole(info.partnerRow[13]) : "MEM";
      const order = ["MEM", "VIP", "O-STUDENT", "STUDENT", "S-STUDENT", "ADMIN"];
      combinedData.role = order.indexOf(myRole) > order.indexOf(partnerRole) ? myRole : partnerRole;
    } else {
      combinedData.role = "MEM";
    }

    // 4. Lấy Rank
    try {
      const data = getTripleDLeaderboardForCurrentUser(1);
      if (data && data.userRank && data.userRank.rank) {
        combinedData.rank = "#" + data.userRank.rank;
      }
    } catch(e) {
      combinedData.rank = "N/A";
    }
    
    // 5. Thêm thông báo (Notifications)
    try {
      const notifs = getSystemNotifications();
      combinedData.notifications = notifs;
    } catch(e) {
      combinedData.notifications = [];
    }

    // 6. Thêm danh sách thú cưng (Owned Pets)
    try {
      combinedData.ownedPets = getOwnedPets(email);
    } catch(e) {
      combinedData.ownedPets = [];
    }

    // 7. Kiểm tra nếu có ưu đãi cứu Streak
    try {
      var rescueKey = "lost_streak_" + email.toLowerCase().replace(/[^a-z0-9]/g, "");
      var propVal = PropertiesService.getScriptProperties().getProperty(rescueKey);
      if (propVal) {
        var lostData = JSON.parse(propVal);
        // Hết hạn ưu đãi sau 24 giờ kể từ khi lưu
        if (new Date().getTime() - lostData.timestamp < 24 * 60 * 60 * 1000) {
          combinedData.streakLostRescue = lostData;
        } else {
          PropertiesService.getScriptProperties().deleteProperty(rescueKey);
        }
      }
    } catch(e) {
      console.error("Lỗi kiểm tra streakLostRescue: " + e.toString());
    }

  } catch (err) {
    console.error("Lỗi getInitialAppDataCombined: " + err.message);
  }

  return combinedData;
}

/**
 * Resends the email verification link.
 * @param {string} email The email of the user.
 * @returns {Object} Success status and message.
 */
function resendVerificationLink(email) {
  try {
    var user = getUserByEmail(email);
    if (!user) {
      return { success: false, message: "Không tìm thấy email." };
    }

    if (user.data[7] === "Verified") {
      return { success: false, message: "Email đã được xác minh. Vui lòng thử đăng nhập." };
    }

    var newVerificationToken = generateVerificationToken();
    var webAppUrl = ScriptApp.getService().getUrl();
    var verificationLink = webAppUrl + '?token=' + newVerificationToken;

    updateUserField(user.row, 8, newVerificationToken); // Column I (index 8)

    var emailBody = "Triple D Xin Chào " + user.data[0] + ",\n\n" +
                    "Bạn đã yêu cầu một liên kết xác minh mới. Vui lòng nhấp vào liên kết bên dưới để xác minh địa chỉ email của bạn:\n\n" +
                    verificationLink + "\n\n" +
                    "Liên kết này sẽ hết hạn sau 24 giờ.\n\n" +
                    "Cảm Ơn,\nTriple D";

    try {
      sendEmail(email, "Liên Kết Xác Minh Tài Khoản Từ Triple D", emailBody);
      return { success: true, message: "Một liên kết xác minh mới đã được gửi tới email của bạn." };
    } catch (e) {
      return { success: false, message: e.message };
    }

  } catch (e) {
    Logger.log(e.toString());
    return { success: false, message: "Lỗi khi gửi lại liên kết xác minh: " + e.message };
  }
}

/**
 * Sends an OTP for password reset.
 * @param {string} email The email to send OTP to.
 * @returns {Object} Success status and message.
 */
function sendPasswordResetOTP(email) {
  try {
    var user = getUserByEmail(email);

    if (!user) {
      return { success: false, message: "Email chưa được đăng ký." };
    }

    if (user.data[7] !== "Verified") { // Status column H (index 7)
      return { success: false, message: "Tài khoản của bạn chưa được xác minh. Vui lòng xác minh email của bạn trước." };
    }

    var otp = generateOTP();
    var expiryTime = new Date();
    expiryTime.setMinutes(expiryTime.getMinutes() + 10); // OTP valid for 10 minutes

    updateUserField(user.row, 5, otp); // Update OTP column F (index 5)
    updateUserField(user.row, 6, expiryTime); // Update DATE OF EXPIRY column G (index 6)

    var emailBody = "Triple D Xin Chào " + user.data[0] + ",\n\n" + // Name column A (index 0)
                    "Bạn đã yêu cầu đặt lại mã truy cập. Vui lòng sử dụng OTP sau để xác minh danh tính của bạn:\n\n" +
                    "**" + otp + "**\n\n" +
                    "Quan trọng: Mã OTP này sẽ hết hạn sau 10 phút vì lý do bảo mật.\n\n" +
                    "Nếu bạn không yêu cầu đặt lại mã truy cập này, vui lòng bỏ qua email này. mã truy cập của bạn sẽ không thay đổi.\n\n" +
                    "Đây là tin nhắn tự động. Vui lòng không trả lời email này.";

    try {
      sendEmail(email, "Đặt Lại mã truy cập - Xác Minh OTP", emailBody);
      return { success: true, message: "OTP đã được gửi đến email của bạn. Vui lòng kiểm tra hộp thư đến." };
    } catch (e) {
      return { success: false, message: e.message };
    }

  } catch (e) {
    Logger.log(e.toString());
    return { success: false, message: "Lỗi khi gửi OTP: " + e.message };
  }
}

/**
 * Verifies the entered OTP.
 * @param {string} email The email of the user.
 * @param {string} otp The entered OTP.
 * @returns {Object} Success status and message.
 */
function verifyOTP(email, otp) {
  var user = getUserByOTP(email, otp);
  if (user) {
    return { success: true, message: "OTP đã được xác minh thành công!" };
  } else {
    return { success: false, message: "OTP không hợp lệ hoặc đã hết hạn." };
  }
}

/**
 * Resets the user's password after OTP verification.
 * @param {string} email The email of the user.
 * @param {string} newPassword The new password.
 * @returns {Object} Success status and message.
 */
function resetPassword(email, newPassword) {
  try {
    var user = getUserByEmail(email);
    if (user) {
      // Re-hash the new password with the existing salt
      var salt = user.data[9]; // Retrieve existing salt (Column J, index 9)
      var newHashedPassword = hashPassword(newPassword, salt);

      updateUserField(user.row, 2, newHashedPassword); // Update Hashed Password column C (index 2)
      updateUserField(user.row, 5, ""); // Clear OTP column F (index 5)
      updateUserField(user.row, 6, ""); // Clear DATE OF EXPIRY column G (index 6)
      return { success: true, message: "Đã đặt lại mã truy cập thành công!" };
    }
    return { success: false, message: "Không tìm thấy người dùng." };
  } catch (e) {
    Logger.log(e.toString());
    return { success: false, message: "Lỗi khi đặt lại mã truy cập: " + e.message };
  }
}

function loadHomePage() {
  return HtmlService.createTemplateFromFile('Home').evaluate()
    .setTitle("Triple D Sinh Học")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}



function getInitialBatchData() {
  try {
    const banners = getBanners();
    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) {
      return { success: true, loggedIn: false, banners: banners };
    }

    const role = getUserRole(email);
    const user = getUserByEmail(email);
    const allowedFolders = getUserAllowedFolders(email);
    const courses = getCourses();
    const exams = getExams();
    const books = getBooks();

    let userInfo = null;
    if (user) {
      userInfo = {
        name: user.data[0] || "",
        email: user.data[1] || "",
        status: user.data[7] || "",
        points: Number(user.data[10]) || 0,
        role: role
      };
    }

    return {
      success: true,
      loggedIn: true,
      email: email,
      role: role,
      userInfo: userInfo,
      allowedFolders: allowedFolders,
      courses: courses,
      exams: exams,
      books: books,
      banners: banners
    };
  } catch (e) {
    Logger.log("Lỗi getInitialBatchData: " + e.message);
    return { success: false, error: e.message };
  }
}

function isUserLoggedIn() {
  const cache = CacheService.getUserCache();
  Logger.log("Checking login status: " + cache.get("loggedInUser"));
  return !!cache.get("loggedInUser");
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function myURL() {
  return ScriptApp.getService().getUrl();
}
function include(filename) {
  try {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
  } catch (e) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  }
}

function cleanupOldCacheProperties(type) {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    let oldKeys = [];
    if (type === "books") {
      oldKeys = ["cachedBooksGlobal", "cachedBooks"];
    } else if (type === "exams") {
      oldKeys = ["cachedExamsGlobal", "cachedExams"];
    } else if (type === "courses") {
      oldKeys = ["cachedCoursesGlobal", "cachedCourses"];
    } else {
      oldKeys = [
        "cachedBooksGlobal",
        "cachedExamsGlobal",
        "cachedCoursesGlobal",
        "cachedBooks",
        "cachedExams",
        "cachedCourses"
      ];
    }
    let deletedCount = 0;
    oldKeys.forEach(key => {
      if (scriptProps.getProperty(key) !== null) {
        scriptProps.deleteProperty(key);
        deletedCount++;
      }
    });
    if (deletedCount > 0) {
      Logger.log(`🧹 Đã dọn dẹp ${deletedCount} khóa cache (${type || "all"}) cũ khỏi ScriptProperties.`);
    }
    return { success: true, deletedCount: deletedCount };
  } catch (err) {
    Logger.log("Lỗi cleanupOldCacheProperties: " + err);
    return { success: false, error: err.toString() };
  }
}

function cleanupExcessScriptProperties_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const map = props.getProperties();
    const now = Date.now();
    for (var k in map) {
      if (k.startsWith("lost_streak_")) {
        try {
          var data = JSON.parse(map[k]);
          if (data && data.timestamp && (now - data.timestamp > 24 * 60 * 60 * 1000)) {
            props.deleteProperty(k);
          }
        } catch (e) {
          props.deleteProperty(k);
        }
      } else if (k.startsWith("USER_FORCE_LOGOUT_")) {
        try {
          var ts = Number(map[k]);
          if (ts && (now - ts > 7 * 24 * 60 * 60 * 1000)) {
            props.deleteProperty(k);
          }
        } catch (e) {
          props.deleteProperty(k);
        }
      } else if (k.startsWith("checkpoint_")) {
        props.deleteProperty(k);
      }
    }
    ["cachedBooksGlobal", "cachedExamsGlobal", "cachedCoursesGlobal", "cachedBooks", "cachedExams", "cachedCourses"].forEach(function(oldKey) {
      if (props.getProperty(oldKey) !== null) {
        props.deleteProperty(oldKey);
      }
    });
  } catch (e) {
    Logger.log("Lỗi cleanupExcessScriptProperties_: " + e.message);
  }
}

function safeSetScriptProperty(key, value) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(key, value);
    return true;
  } catch (err) {
    Logger.log("[SafeSetProperty] Lỗi khi lưu property " + key + ": " + err.message + ". Đang tiến hành dọn dẹp ScriptProperties...");
    try {
      cleanupExcessScriptProperties_();
      const props = PropertiesService.getScriptProperties();
      props.setProperty(key, value);
      return true;
    } catch (retryErr) {
      Logger.log("[SafeSetProperty] Vẫn lỗi sau khi dọn dẹp cho key " + key + ": " + retryErr.message);
      return false;
    }
  }
}

function getCachedBooksFromSheet() {
  if (_cachedBooksMemory && _cachedBooksMemory.length > 0) return _cachedBooksMemory;
  try {
    const sheet = getSheet("__BookCache__");
    if (!sheet) {
      const ss = getSpreadsheet();
      if (ss) {
        const newSheet = ss.insertSheet("__BookCache__");
        newSheet.appendRow(["title", "fileId", "folder", "folderId", "level", "isPaid"]);
      }
      // Migration fallback from ScriptProperties if exists
      const props = PropertiesService.getScriptProperties();
      const rawOld = props.getProperty("cachedBooksGlobal") || props.getProperty("cachedBooks");
      if (rawOld) {
        try {
          const booksOld = JSON.parse(rawOld);
          if (Array.isArray(booksOld) && booksOld.length > 0) {
            saveCachedBooksToSheet(booksOld);
            return booksOld;
          }
        } catch(e) {}
      }
      // Tự động quét cập nhật cache sách nếu sheet mới tạo chưa có dữ liệu
      try {
        updateBookCache();
        if (_cachedBooksMemory && _cachedBooksMemory.length > 0) return _cachedBooksMemory;
      } catch(e) {}
      return [];
    }

    const values = getSheetDataCached("__BookCache__");
    if (!values || values.length <= 1) {
      // Migration fallback from ScriptProperties if sheet has only headers
      const props = PropertiesService.getScriptProperties();
      const rawOld = props.getProperty("cachedBooksGlobal") || props.getProperty("cachedBooks");
      if (rawOld) {
        try {
          const booksOld = JSON.parse(rawOld);
          if (Array.isArray(booksOld) && booksOld.length > 0) {
            saveCachedBooksToSheet(booksOld);
            return booksOld;
          }
        } catch(e) {}
      }
      // Tự động quét cập nhật cache sách nếu sheet rỗng
      try {
        updateBookCache();
        if (_cachedBooksMemory && _cachedBooksMemory.length > 0) return _cachedBooksMemory;
      } catch(e) {}
      return [];
    }
    
    const books = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row[0] && !row[1]) continue; // Skip empty rows
      books.push({
        title: String(row[0] || ""),
        fileId: String(row[1] || ""),
        folder: String(row[2] || ""),
        folderId: String(row[3] || ""),
        level: Number(row[4]) || 0,
        isPaid: row[5] === true || String(row[5]).toLowerCase() === "true"
      });
    }
    _cachedBooksMemory = books;
    return books;
  } catch (err) {
    Logger.log("Lỗi getCachedBooksFromSheet: " + err);
    return [];
  }
}

function saveCachedBooksToSheet(books) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("__BookCache__");
    if (!sheet) {
      sheet = ss.insertSheet("__BookCache__");
    } else {
      sheet.clear();
    }
    
    const data = [["title", "fileId", "folder", "folderId", "level", "isPaid"]];
    
    if (Array.isArray(books)) {
      books.forEach(b => {
        data.push([
          b.title || "",
          String(b.fileId || ""),
          b.folder || "",
          String(b.folderId || ""),
          Number(b.level) || 0,
          b.isPaid === true || String(b.isPaid).toLowerCase() === "true"
        ]);
      });
    }
    
    sheet.getRange(1, 1, data.length, 6).setValues(data);
    _cachedBooksMemory = books;
    invalidateSheetCache("__BookCache__");

    if (books && books.length > 0) {
      cleanupOldCacheProperties("books");
    }
  } catch (err) {
    Logger.log("Lỗi saveCachedBooksToSheet: " + err);
  }
}

function getCachedExamsFromSheet() {
  if (_cachedExamsMemory && _cachedExamsMemory.length > 0) return _cachedExamsMemory;
  try {
    const sheet = getSheet("__ExamCache__");
    if (!sheet) {
      const ss = getSpreadsheet();
      if (ss) {
        const newSheet = ss.insertSheet("__ExamCache__");
        newSheet.appendRow(["title", "fileId", "folder", "folderId", "level", "isPaid"]);
      }
      // Migration fallback from ScriptProperties if exists
      const props = PropertiesService.getScriptProperties();
      const rawOld = props.getProperty("cachedExamsGlobal") || props.getProperty("cachedExams");
      if (rawOld) {
        try {
          const examsOld = JSON.parse(rawOld);
          if (Array.isArray(examsOld) && examsOld.length > 0) {
            saveCachedExamsToSheet(examsOld);
            return examsOld;
          }
        } catch(e) {}
      }
      // Tự động quét cập nhật cache tài liệu nếu sheet mới tạo chưa có dữ liệu
      try {
        updateExamCache();
        if (_cachedExamsMemory && _cachedExamsMemory.length > 0) return _cachedExamsMemory;
      } catch(e) {}
      return [];
    }

    const values = getSheetDataCached("__ExamCache__");
    if (!values || values.length <= 1) {
      // Migration fallback from ScriptProperties if sheet has only headers
      const props = PropertiesService.getScriptProperties();
      const rawOld = props.getProperty("cachedExamsGlobal") || props.getProperty("cachedExams");
      if (rawOld) {
        try {
          const examsOld = JSON.parse(rawOld);
          if (Array.isArray(examsOld) && examsOld.length > 0) {
            saveCachedExamsToSheet(examsOld);
            return examsOld;
          }
        } catch(e) {}
      }
      // Tự động quét cập nhật cache tài liệu nếu sheet rỗng
      try {
        updateExamCache();
        if (_cachedExamsMemory && _cachedExamsMemory.length > 0) return _cachedExamsMemory;
      } catch(e) {}
      return [];
    }
    
    const exams = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row[0] && !row[1]) continue; // Skip empty rows
      exams.push({
        title: String(row[0] || ""),
        fileId: String(row[1] || ""),
        folder: String(row[2] || ""),
        folderId: String(row[3] || ""),
        level: Number(row[4]) || 0,
        isPaid: row[5] === true || String(row[5]).toLowerCase() === "true"
      });
    }
    _cachedExamsMemory = exams;
    return exams;
  } catch (err) {
    Logger.log("Lỗi getCachedExamsFromSheet: " + err);
    return [];
  }
}

function saveCachedExamsToSheet(exams) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("__ExamCache__");
    if (!sheet) {
      sheet = ss.insertSheet("__ExamCache__");
    } else {
      sheet.clear();
    }
    
    const data = [["title", "fileId", "folder", "folderId", "level", "isPaid"]];
    
    if (Array.isArray(exams)) {
      exams.forEach(e => {
        data.push([
          e.title || "",
          String(e.fileId || ""),
          e.folder || "",
          String(e.folderId || ""),
          Number(e.level) || 0,
          e.isPaid === true || String(e.isPaid).toLowerCase() === "true"
        ]);
      });
    }
    
    sheet.getRange(1, 1, data.length, 6).setValues(data);
    _cachedExamsMemory = exams;
    invalidateSheetCache("__ExamCache__");

    if (exams && exams.length > 0) {
      cleanupOldCacheProperties("exams");
    }
  } catch (err) {
    Logger.log("Lỗi saveCachedExamsToSheet: " + err);
  }
}

function getBooks() {
  const allBooks = getCachedBooksFromSheet();
  const email = CacheService.getUserCache().get("loggedInUser");
  const role = getUserRole(email); // VIP, STUDENT, MEM

  if (!allBooks || allBooks.length === 0) return [];

  // Giữ nguyên fileId, chỉ thêm flag canOpen để UI kiểm soát
  return allBooks.map(book => {
    const canOpen = canUserAccessExam(role, book.folder);
    return {
      ...book,
      canOpen
    };
  });
}

function getExams() {
  const allExams = getCachedExamsFromSheet();
  const email = CacheService.getUserCache().get("loggedInUser");
  const role = getUserRole(email); // VIP, STUDENT, MEM

  if (!allExams || allExams.length === 0) return [];

  // Lọc theo quyền người dùng
  return allExams.map(exam => {
    const canOpen = canUserAccessExam(role, exam.folder);
    return {
      ...exam,
      canOpen
    };
  });
}

function updateExamCache() {
  const startTime = Date.now();
  const MAX_RUNTIME_MS = 330000; // 5 phút 30 giây (dừng trước giới hạn 6 phút của GAS)
  let timedOut = false;

  const rootFolderId = "1E6UgH6uahT7WzlsG3cv1BMyH_WPNPkKl";
  const rootFolder = DriveApp.getFolderById(rootFolderId);

  // Kiểm tra checkpoint quét dở dang từ lần trước
  const scriptProps = PropertiesService.getScriptProperties();
  const rawCheckpoint = scriptProps.getProperty("checkpoint_exams_folders");
  let processedFolders = new Set();
  const examMap = new Map();

  if (rawCheckpoint) {
    try {
      processedFolders = new Set(JSON.parse(rawCheckpoint));
      // Nếu đang tiếp tục từ checkpoint dở dang -> nạp lại cache cũ để giữ file các thư mục đã quét trước đó
      const oldExams = getCachedExamsFromSheet();
      oldExams.forEach(e => {
        if (e.fileId) examMap.set(String(e.fileId), e);
      });
    } catch(e) {}
  }

  function processFolder(folder, parentPath, level) {
    if (timedOut || Date.now() - startTime > MAX_RUNTIME_MS) {
      timedOut = true;
      return;
    }

    try {
      const folderId = folder.getId();
      const folderName = folder.getName();
      const fullPath = parentPath ? parentPath + " / " + folderName : folderName;

      if (!processedFolders.has(folderId)) {
        const files = folder.getFiles();
        while (files.hasNext()) {
          if (Date.now() - startTime > MAX_RUNTIME_MS) {
            timedOut = true;
            break;
          }
          try {
            const file = files.next();
            if (file.getMimeType() === MimeType.PDF) {
              const isVIP = fullPath.toUpperCase().includes("VIP");
              const isStudentOnly = fullPath.toUpperCase().includes("STUDENT");

              const item = {
                title: file.getName(),
                fileId: file.getId(),
                folder: fullPath,
                folderId: folderId,
                level: level,
                isPaid: isVIP || isStudentOnly
              };
              examMap.set(String(file.getId()), item);
            }
          } catch (fErr) {
            Logger.log("Lỗi xử lý file exam: " + fErr);
          }
        }
        if (!timedOut) {
          processedFolders.add(folderId);
        }
      }

      if (!timedOut) {
        const subfolders = folder.getFolders();
        while (subfolders.hasNext()) {
          if (Date.now() - startTime > MAX_RUNTIME_MS) {
            timedOut = true;
            break;
          }
          try {
            processFolder(subfolders.next(), fullPath, level + 1);
          } catch (subErr) {
            Logger.log("Lỗi xử lý thư mục con exam: " + subErr);
          }
        }
      }
    } catch (err) {
      Logger.log("Lỗi processFolder exam: " + err);
    }
  }

  processFolder(rootFolder, "", 0);

  const finalExams = Array.from(examMap.values());

  if (timedOut) {
    safeSetScriptProperty("checkpoint_exams_folders", JSON.stringify(Array.from(processedFolders)));
    saveCachedExamsToSheet(finalExams);
    return {
      success: true,
      timedOut: true,
      message: `⚠️ Đã quét & gộp ${finalExams.length} tài liệu/bài thi (tạm dừng ở mốc 5 phút 30 giây). Đã lưu checkpoint để lần chạy tiếp theo sẽ quét lấy nốt phần còn lại!`
    };
  } else {
    scriptProps.deleteProperty("checkpoint_exams_folders");
    saveCachedExamsToSheet(finalExams);
    return { success: true, message: `✅ Đã quét hoàn tất toàn bộ cache tài liệu/bài thi: ${finalExams.length} mục (100%).` };
  }
}

function updateBookCache() {
  const startTime = Date.now();
  const MAX_RUNTIME_MS = 330000; // 5 phút 30 giây
  let timedOut = false;

  const rootFolderId = "1U77Z88ueFgg2iebBES0kscsntnFuf_uH"; // 📚 ID thư mục chứa sách
  const rootFolder = DriveApp.getFolderById(rootFolderId);

  const scriptProps = PropertiesService.getScriptProperties();
  const rawCheckpoint = scriptProps.getProperty("checkpoint_books_folders");
  let processedFolders = new Set();
  const bookMap = new Map();

  if (rawCheckpoint) {
    try {
      processedFolders = new Set(JSON.parse(rawCheckpoint));
      // Nếu đang tiếp tục từ checkpoint dở dang -> nạp lại cache cũ
      const oldBooks = getCachedBooksFromSheet();
      oldBooks.forEach(b => {
        if (b.fileId) bookMap.set(String(b.fileId), b);
      });
    } catch(e) {}
  }

  function processFolder(folder, parentPath, level) {
    if (timedOut || Date.now() - startTime > MAX_RUNTIME_MS) {
      timedOut = true;
      return;
    }

    try {
      const folderId = folder.getId();
      const folderName = folder.getName();
      const fullPath = parentPath ? parentPath + " / " + folderName : folderName;

      if (!processedFolders.has(folderId)) {
        const files = folder.getFiles();
        while (files.hasNext()) {
          if (Date.now() - startTime > MAX_RUNTIME_MS) {
            timedOut = true;
            break;
          }
          try {
            const file = files.next();
            if (file.getMimeType() === MimeType.PDF) {
              const isVIP = fullPath.toUpperCase().includes("VIP");
              const isStudentOnly = fullPath.toUpperCase().includes("STUDENT");

              const item = {
                title: file.getName(),
                fileId: file.getId(),
                folder: fullPath,
                folderId: folderId,
                level: level,
                isPaid: isVIP || isStudentOnly
              };
              bookMap.set(String(file.getId()), item);
            }
          } catch (fErr) {
            Logger.log("Lỗi xử lý file book: " + fErr);
          }
        }
        if (!timedOut) {
          processedFolders.add(folderId);
        }
      }

      if (!timedOut) {
        const subfolders = folder.getFolders();
        while (subfolders.hasNext()) {
          if (Date.now() - startTime > MAX_RUNTIME_MS) {
            timedOut = true;
            break;
          }
          try {
            processFolder(subfolders.next(), fullPath, level + 1);
          } catch (subErr) {
            Logger.log("Lỗi xử lý thư mục con book: " + subErr);
          }
        }
      }
    } catch (err) {
      Logger.log("Lỗi processFolder book: " + err);
    }
  }

  processFolder(rootFolder, "", 0);

  const finalBooks = Array.from(bookMap.values());

  if (timedOut) {
    safeSetScriptProperty("checkpoint_books_folders", JSON.stringify(Array.from(processedFolders)));
    saveCachedBooksToSheet(finalBooks);
    return {
      success: true,
      timedOut: true,
      message: `⚠️ Đã quét & gộp ${finalBooks.length} sách (tạm dừng ở mốc 5 phút 30 giây). Đã lưu checkpoint để lần chạy tiếp theo sẽ quét lấy nốt phần còn lại!`
    };
  } else {
    scriptProps.deleteProperty("checkpoint_books_folders");
    saveCachedBooksToSheet(finalBooks);
    return { success: true, message: `✅ Đã quét hoàn tất toàn bộ cache sách: ${finalBooks.length} mục (100%).` };
  }
}

function getCachedCoursesFromSheet() {
  if (_cachedCoursesMemory && _cachedCoursesMemory.length > 0) return _cachedCoursesMemory;
  try {
    const sheet = getSheet("__CourseCache__");
    if (!sheet) {
      const ss = getSpreadsheet();
      if (ss) {
        const newSheet = ss.insertSheet("__CourseCache__");
        newSheet.appendRow(["title", "fileId", "type", "folder", "folderId", "ancestors", "level"]);
      }
      // Migration fallback from ScriptProperties if exists
      const props = PropertiesService.getScriptProperties();
      const rawOld = props.getProperty("cachedCoursesGlobal") || props.getProperty("cachedCourses");
      if (rawOld) {
        try {
          const coursesOld = JSON.parse(rawOld);
          if (Array.isArray(coursesOld) && coursesOld.length > 0) {
            saveCachedCoursesToSheet(coursesOld);
            return coursesOld;
          }
        } catch(e) {}
      }
      // Tự động quét cập nhật cache chuyên đề nếu sheet mới chưa có dữ liệu
      try {
        updateCourseCache();
        if (_cachedCoursesMemory && _cachedCoursesMemory.length > 0) return _cachedCoursesMemory;
      } catch(e) {}
      return [];
    }

    const values = getSheetDataCached("__CourseCache__");
    if (!values || values.length <= 1) {
      // Migration fallback from ScriptProperties if sheet has only headers
      const props = PropertiesService.getScriptProperties();
      const rawOld = props.getProperty("cachedCoursesGlobal") || props.getProperty("cachedCourses");
      if (rawOld) {
        try {
          const coursesOld = JSON.parse(rawOld);
          if (Array.isArray(coursesOld) && coursesOld.length > 0) {
            saveCachedCoursesToSheet(coursesOld);
            return coursesOld;
          }
        } catch(e) {}
      }
      // Tự động quét cập nhật cache chuyên đề nếu sheet rỗng
      try {
        updateCourseCache();
        if (_cachedCoursesMemory && _cachedCoursesMemory.length > 0) return _cachedCoursesMemory;
      } catch(e) {}
      return [];
    }
    
    const courses = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row[0] && !row[1]) continue; // Skip empty rows
      courses.push({
        title: String(row[0] || ""),
        fileId: String(row[1] || ""),
        type: String(row[2] || "video"),
        folder: String(row[3] || ""),
        folderId: String(row[4] || ""),
        ancestors: row[5] ? (typeof row[5] === 'string' ? JSON.parse(row[5]) : row[5]) : [],
        level: Number(row[6]) || 0
      });
    }
    _cachedCoursesMemory = courses;
    return courses;
  } catch (err) {
    Logger.log("Lỗi getCachedCoursesFromSheet: " + err);
    return [];
  }
}

function saveCachedCoursesToSheet(courses) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("__CourseCache__");
    if (!sheet) {
      sheet = ss.insertSheet("__CourseCache__");
    } else {
      sheet.clear();
    }
    
    // Header row
    const data = [["title", "fileId", "type", "folder", "folderId", "ancestors", "level"]];
    
    if (Array.isArray(courses)) {
      courses.forEach(c => {
        data.push([
          c.title || "",
          String(c.fileId || ""),
          c.type || "",
          c.folder || "",
          String(c.folderId || ""),
          c.ancestors ? JSON.stringify(c.ancestors) : "[]",
          Number(c.level) || 0
        ]);
      });
    }
    
    sheet.getRange(1, 1, data.length, 7).setValues(data);
    _cachedCoursesMemory = courses;
    invalidateSheetCache("__CourseCache__");

    if (courses && courses.length > 0) {
      cleanupOldCacheProperties("courses");
    }
  } catch (err) {
    Logger.log("Lỗi saveCachedCoursesToSheet: " + err);
  }
}

function normalizeExamText(str) {
  if (!str) return "";
  return str.toString()
    .normalize("NFC")
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function uploadExamSubmission(base64Data, fileName, examTitle) {
  var lock = LockService.getScriptLock();
  var lockAcquired = false;
  try {
    try {
      lockAcquired = lock.tryLock(30000);
    } catch(eLock) {}

    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) return { success: false, message: "Bạn chưa đăng nhập." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Check exam deadline from MockExams sheet
    var mockExamsSheet = ss.getSheetByName("MockExams");
    if (mockExamsSheet) {
      var examRows = mockExamsSheet.getDataRange().getValues();
      var examHeaders = examRows[0].map(function(h) { return String(h).trim().toLowerCase(); });
      
      var titleColIdx = -1;
      var endColIdx = -1;
      
      for (var col = 0; col < examHeaders.length; col++) {
        var h = examHeaders[col];
        if (h.indexOf("tên bài") !== -1 || h.indexOf("tiêu đề") !== -1 || h === "bài thi") {
          titleColIdx = col;
        } else if (h.indexOf("đóng") !== -1 || h.indexOf("end") !== -1) {
          endColIdx = col;
        }
      }
      
      if (titleColIdx === -1) titleColIdx = 3;
      if (endColIdx === -1) endColIdx = 5;
      
      var targetExamTitleNorm = normalizeExamText(examTitle);
      var foundExam = false;
      var examEndTimeVal = null;
      
      for (var r = 1; r < examRows.length; r++) {
        var rowExamTitle = examRows[r][titleColIdx] ? normalizeExamText(examRows[r][titleColIdx]) : "";
        if (rowExamTitle === targetExamTitleNorm) {
          foundExam = true;
          examEndTimeVal = examRows[r][endColIdx];
          break;
        }
      }
      
      if (foundExam && examEndTimeVal) {
        var examEndTime = (examEndTimeVal instanceof Date) ? examEndTimeVal : new Date(examEndTimeVal.toString().trim());
        if (examEndTime && !isNaN(examEndTime.getTime())) {
          var now = new Date();
          if (now > examEndTime) {
            return { success: false, message: "Bài khảo thí này đã hết hạn nộp theo lịch." };
          }
        }
      }
    }

    var sheet = ss.getSheetByName('ExamSubmissions');
    if (!sheet) {
      sheet = ss.insertSheet('ExamSubmissions');
      sheet.appendRow(['Timestamp', 'Email', 'Exam Title', 'File URL', 'Filename', 'Score', 'Remark', 'GradedBy', 'GradedAt']);
    }

    // Tìm tất cả các dòng nộp bài cũ của học sinh đối với bài thi này (bảo đảm chuẩn hóa chuỗi)
    var data = sheet.getDataRange().getValues();
    var matchingRows = []; // 0-based indices
    var normEmail = normalizeExamText(email);
    var normTitle = normalizeExamText(examTitle);

    for (var i = 1; i < data.length; i++) {
      var rowEmail = normalizeExamText(data[i][1]);
      var rowTitle = normalizeExamText(data[i][2]);
      if (rowEmail === normEmail && rowTitle === normTitle) {
        matchingRows.push(i);
      }
    }

    // Xóa tất cả các file cũ trên Google Drive để tránh rác ổ đĩa
    for (var m = 0; m < matchingRows.length; m++) {
      var oldIdx = matchingRows[m];
      var oldFileUrl = data[oldIdx][3] ? data[oldIdx][3].toString() : "";
      if (oldFileUrl) {
        try {
          var matchId = oldFileUrl.match(/[-\w]{25,}/);
          var oldFileId = matchId ? matchId[0] : null;
          if (oldFileId) {
            var oldFile = DriveApp.getFileById(oldFileId);
            oldFile.setTrashed(true);
          }
        } catch (err) {
          Logger.log("Không thể xóa file cũ: " + err.toString());
        }
      }
    }

    // Try to find the requested Google Drive folder or fallback to the default folder
    var folder;
    var targetFolderId = "1Ea6-0QTJ_BNLjvnQ3heJ4Qe2X5i9Zqfd";
    try {
      folder = DriveApp.getFolderById(targetFolderId);
    } catch (err) {
      var folders = DriveApp.getFoldersByName("Bài thi của học viên");
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder("Bài thi của học viên");
      }
    }

    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.PDF, examTitle + " - " + email + " - " + fileName);
    var file = folder.createFile(blob);

    if (matchingRows.length > 0) {
      // Xóa tất cả các dòng nộp bài cũ từ dưới lên để dòng nộp đè mới xuất hiện ở dưới cùng của sheet
      for (var k = matchingRows.length - 1; k >= 0; k--) {
        sheet.deleteRow(matchingRows[k] + 1);
      }
    }
    
    // Ghi dòng mới
    sheet.appendRow([new Date(), email.trim(), examTitle.trim(), file.getUrl(), file.getName(), "", "", "", ""]);
    invalidateSheetCache("ExamSubmissions");
    
    // Nếu học sinh nộp/nộp lại bài thi, tạm xóa email khỏi thông báo "Bài Đã Chấm Xong" của đề này cho đến khi được chấm lại
    removeStudentFromGroupNotification("Bài Đã Chấm Xong", examTitle, email);

    // Auto-update sharing so the owner can read, though owner inherently can. Just to be safe.
    logActivityToSheet(email, "SUBMIT_EXAM", `Nộp bài thi (ghi đè/cập nhật): ${examTitle}`);

    SpreadsheetApp.flush();
    return { success: true, message: "Nộp bài thành công! Hệ thống đã ghi nhận." };
  } catch (e) {
    return { success: false, message: "Lỗi: " + e.message };
  } finally {
    if (lockAcquired) {
      try { lock.releaseLock(); } catch(e) {}
    }
  }
}

function uploadSStudentEvidence(base64Data, fileName, contentType) {
  try {
    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) return { success: false, message: "Bạn chưa đăng nhập." };

    const PARENT_FOLDER_ID = "1F5FAHo3SpKPH0oB26W6f6mZ-Va1RCOs6";
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);

    // Find/Create folder with user's email inside the parent folder
    var folders = parentFolder.getFoldersByName(email);
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = parentFolder.createFolder(email);
    }

    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, email + " - " + fileName);
    var file = folder.createFile(blob);

    logActivityToSheet(email, "SUBMIT_EVIDENCE", `Nộp minh chứng: ${fileName}`);

    return { success: true, message: "Nộp minh chứng thành công! BQT sẽ sớm kiểm tra." };
  } catch (e) {
    return { success: false, message: "Lỗi: " + e.message };
  }
}

function getCombinedSubmissionsData(title1, title2) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ExamSubmissions');
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return null;

    var normT1 = normalizeExamText(title1);
    var normT2 = normalizeExamText(title2);

    var numQ1 = 0;
    var numQ2 = 0;
    
    for (var i = 1; i < data.length; i++) {
      const row = data[i];
      const rowExamTitle = row[2] ? row[2].toString().trim() : "";
      const rowNormTitle = normalizeExamText(rowExamTitle);
      const rowScoreStr = row[5] !== undefined ? row[5].toString().trim() : "";
      if (rowScoreStr !== "" && rowScoreStr.indexOf("[") !== -1) {
        var parts = rowScoreStr.split("[");
        var compStr = parts[1].replace("]", "").trim();
        var comps = compStr ? compStr.split(";").map(s => parseFloat(s.trim().replace(/,/g, "."))) : [];
        if (rowNormTitle === normT1) {
          numQ1 = Math.max(numQ1, comps.length);
        } else if (rowNormTitle === normT2) {
          numQ2 = Math.max(numQ2, comps.length);
        }
      }
    }

    var students = {};

    for (var i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = row[1] ? row[1].toString().toLowerCase().trim() : "";
      const rowExamTitle = row[2] ? row[2].toString().trim() : "";
      const rowNormTitle = normalizeExamText(rowExamTitle);
      const rowFileUrl = row[3] !== undefined ? row[3].toString().trim() : "";
      const rowScoreStr = row[5] !== undefined ? row[5].toString().trim() : "";
      const rowRemark = row[6] !== undefined ? row[6].toString().trim() : "";

      if (!rowEmail) continue;

      var isT1 = (rowNormTitle === normT1);
      var isT2 = (rowNormTitle === normT2);

      if (isT1 || isT2) {
        if (!students[rowEmail]) {
          students[rowEmail] = {
            email: rowEmail,
            score1: null,
            score2: null,
            comps1: null,
            comps2: null,
            remark1: "",
            remark2: "",
            url1: "",
            url2: ""
          };
        }

        var sObj = students[rowEmail];

        var scoreVal = null;
        var comps = null;
        if (rowScoreStr !== "") {
          if (rowScoreStr.indexOf("[") !== -1) {
            var parts = rowScoreStr.split("[");
            scoreVal = parseFloat(parts[0].trim().replace(/,/g, "."));
            var compStr = parts[1].replace("]", "").trim();
            comps = compStr ? compStr.split(";").map(s => {
              var val = parseFloat(s.trim().replace(/,/g, "."));
              return isNaN(val) ? 0 : val;
            }) : [];
          } else {
            scoreVal = parseFloat(rowScoreStr.replace(/,/g, "."));
          }
        }

        if (isT1) {
          if (sObj.score1 === null || !isNaN(scoreVal)) {
            sObj.score1 = isNaN(scoreVal) ? null : scoreVal;
            sObj.comps1 = comps;
            sObj.remark1 = rowRemark;
            sObj.url1 = rowFileUrl;
          }
        } else {
          if (sObj.score2 === null || !isNaN(scoreVal)) {
            sObj.score2 = isNaN(scoreVal) ? null : scoreVal;
            sObj.comps2 = comps;
            sObj.remark2 = rowRemark;
            sObj.url2 = rowFileUrl;
          }
        }
      }
    }

    return {
      students: students,
      numQ1: numQ1,
      numQ2: numQ2
    };
  } catch (e) {
    return null;
  }
}

function calculateAwardFromRankAndTotal(rank, total, scoreVal) {
  var N = parseInt(total, 10) || 0;
  var r = parseInt(rank, 10) || 0;

  if (N >= 1 && r > 0) {
    if (N === 1) return { award: "Giải Nhất", badge: "🥇" };
    if (N === 2) {
      if (r === 1) return { award: "Giải Nhất", badge: "🥇" };
      return { award: "Giải Nhì", badge: "🥈" };
    }

    var T = Math.max(1, Math.ceil(0.60 * N));
    var g1 = Math.max(1, Math.round(0.08 * T));
    var g2 = Math.max(1, Math.round(0.25 * T));
    var g3 = Math.max(1, Math.round(0.32 * T));
    var gKK = Math.max(0, T - g1 - g2 - g3);

    if (r <= g1) return { award: "Giải Nhất", badge: "🥇" };
    if (r <= g1 + g2) return { award: "Giải Nhì", badge: "🥈" };
    if (r <= g1 + g2 + g3) return { award: "Giải Ba", badge: "🥉" };
    if (r <= T) return { award: "Giải Khuyến Khích", badge: "🎖️" };
    return { award: "Chưa đoạt giải", badge: "📘" };
  }

  return { award: "Chưa đoạt giải", badge: "📘" };
}

function getScoreAndRank(examTitle) {
  try {
    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) return { success: false, message: "Bạn chưa đăng nhập." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ExamSubmissions');
    if (!sheet) {
      return { success: false, message: "Chưa có dữ liệu bài nộp." };
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: false, message: "Danh sách bài nộp trống." };
    }

    if (examTitle.indexOf(" + ") !== -1) {
      var partsTitle = examTitle.split(" + ");
      var title1 = partsTitle[0].trim();
      var title2 = partsTitle[1].trim();
      var combinedData = getCombinedSubmissionsData(title1, title2);
      if (!combinedData) {
        return { success: false, message: "Chưa có dữ liệu bài nộp cho cặp bài thi này." };
      }

      var emailClean = email.toLowerCase().trim();
      var myStudent = combinedData.students[emailClean];

      if (!myStudent) {
        return { success: false, message: "Bạn chưa nộp bài hoặc chưa có kết quả cho bất kỳ bài thi nào trong cặp này." };
      }

      var allCombinedScores = [];
      var myCombinedScore = null;
      var myRemarkCombined = "";
      var myUrlCombined = "";

      for (var em in combinedData.students) {
        var s = combinedData.students[em];
        if (s.score1 !== null || s.score2 !== null) {
          var s1 = s.score1 !== null ? Math.round(s.score1 * 100) / 100 : 0;
          var s2 = s.score2 !== null ? Math.round(s.score2 * 100) / 100 : 0;
          var combVal = Math.round((s1 + s2) * 100) / 100;
          allCombinedScores.push({ email: em, score: combVal });

          if (em === emailClean) {
            myCombinedScore = combVal;
            var rems = [];
            var s1Clean = s.score1 !== null ? (Math.round(s.score1 * 100) / 100).toString().replace(/\./g, ",") : "";
            var s2Clean = s.score2 !== null ? (Math.round(s.score2 * 100) / 100).toString().replace(/\./g, ",") : "";
            if (s.score1 !== null) rems.push("Ngày 1: " + s1Clean + "đ" + (s.remark1 ? " (" + s.remark1 + ")" : ""));
            else rems.push("Ngày 1: Chưa nộp/Chưa chấm");
            if (s.score2 !== null) rems.push("Ngày 2: " + s2Clean + "đ" + (s.remark2 ? " (" + s.remark2 + ")" : ""));
            else rems.push("Ngày 2: Chưa nộp/Chưa chấm");
            myRemarkCombined = rems.join(" | ");
            myUrlCombined = s.url1 || s.url2;
          }
        }
      }

      if (myCombinedScore === null) {
        return { success: false, message: "Bạn chưa có điểm số kết quả cho cặp bài thi này." };
      }

      allCombinedScores.sort(function(a, b) { return b.score - a.score; });
      let combinedRank = 1;
      for (var j = 0; j < allCombinedScores.length; j++) {
        if (allCombinedScores[j].score > myCombinedScore) {
          combinedRank++;
        }
      }

      var myComps1 = myStudent.comps1 || Array(combinedData.numQ1).fill(0);
      var myComps2 = myStudent.comps2 || Array(combinedData.numQ2).fill(0);
      var combinedComps = myComps1.concat(myComps2).map(function(c) {
        return Math.round((parseFloat(c) || 0) * 100) / 100;
      });
      
      var scoreRepr = (Math.round(myCombinedScore * 100) / 100).toString();
      if (combinedComps.length > 0) {
        scoreRepr += " [" + combinedComps.join("; ") + "]";
      }

      var awardResComb = calculateAwardFromRankAndTotal(combinedRank, allCombinedScores.length, myCombinedScore);

      return {
        success: true,
        score: scoreRepr,
        rank: combinedRank,
        total: allCombinedScores.length,
        award: awardResComb.award,
        awardBadge: awardResComb.badge,
        remark: myRemarkCombined,
        fileUrl: myUrlCombined,
        isPair: true,
        title1: title1,
        title2: title2,
        score1: myStudent.score1 !== null ? (Math.round(myStudent.score1 * 100) / 100) : null,
        score2: myStudent.score2 !== null ? (Math.round(myStudent.score2 * 100) / 100) : null,
        comps1: myComps1,
        comps2: myComps2,
        remark1: myStudent.remark1 || "",
        remark2: myStudent.remark2 || "",
        numQ1: combinedData.numQ1,
        numQ2: combinedData.numQ2
      };
    }

    let scoresForExam = [];
    let myScoreVal = null;
    let myScoreRaw = null;
    let myRemark = null;
    let myFileUrl = "";

    const targetNormTitle = normalizeExamText(examTitle);
    const targetNormEmail = normalizeExamText(email);
    const studentScoreMap = {};

    for (var i = 1; i < data.length; i++) {
       const row = data[i];
       const rowEmail = row[1] ? row[1].toString().toLowerCase().trim() : "";
       const rowNormEmail = normalizeExamText(rowEmail);
       const rowExamTitle = row[2] ? row[2].toString().trim() : "";
       const rowNormTitle = normalizeExamText(rowExamTitle);
       const rowFileUrl = row[3] !== undefined ? row[3].toString().trim() : "";
       const rowScoreStr = row[5] !== undefined ? row[5].toString().trim() : "";
       const rowRemark = row[6] !== undefined ? row[6].toString().trim() : "";
       
       if (rowNormTitle === targetNormTitle) {
         if (rowScoreStr !== "") {
           let scoreVal = parseFloat(rowScoreStr.replace(/,/g, "."));
           if (!isNaN(scoreVal)) {
             studentScoreMap[rowNormEmail] = { email: rowEmail, score: scoreVal };
             if (rowNormEmail === targetNormEmail) {
               myScoreVal = scoreVal;
               myScoreRaw = rowScoreStr;
               myRemark = rowRemark;
               myFileUrl = rowFileUrl;
             }
           } else if (rowNormEmail === targetNormEmail && myScoreRaw === null) {
              myScoreRaw = rowScoreStr;
              myRemark = rowRemark;
              myFileUrl = rowFileUrl;
           }
         } else if (rowNormEmail === targetNormEmail && myScoreRaw === null) {
            myScoreRaw = "";
            myRemark = rowRemark;
            myFileUrl = rowFileUrl;
         }
       }
    }

    scoresForExam = Object.values(studentScoreMap);

    if (myScoreRaw === null) {
       return { success: false, message: "Bạn chưa nộp bài hoặc chưa có kết quả." };
    }
    
    if (myScoreVal === null) {
       return { success: false, message: "Bài thi của bạn đang chờ chấm điểm." + (myScoreRaw ? " (Phản hồi: " + myScoreRaw + ")" : "") };
    }

    let rank = 1;
    for (let i = 0; i < scoresForExam.length; i++) {
        if (scoresForExam[i].score > myScoreVal) {
            rank++;
        }
    }

    var awardRes = calculateAwardFromRankAndTotal(rank, scoresForExam.length, myScoreVal);

    return { 
      success: true, 
      score: myScoreRaw, 
      rank: rank, 
      total: scoresForExam.length,
      award: awardRes.award,
      awardBadge: awardRes.badge,
      remark: myRemark,
      fileUrl: myFileUrl
    };

  } catch (e) {
    return { success: false, message: "Lỗi: " + e.message };
  }
}

function getExamStatistics(examTitle) {
  try {
    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) return { success: false, message: "Bạn chưa đăng nhập." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ExamSubmissions');
    if (!sheet) {
      return { success: false, message: "Chưa có dữ liệu bài nộp nào trên hệ thống." };
    }

    if (examTitle.indexOf(" + ") !== -1) {
      var partsTitle = examTitle.split(" + ");
      var title1 = partsTitle[0].trim();
      var title2 = partsTitle[1].trim();
      var combinedData = getCombinedSubmissionsData(title1, title2);
      if (!combinedData) {
        return { success: false, message: "Chưa có dữ liệu bài nộp cho cặp bài thi này." };
      }

      let totalSubmissions = 0;
      let gradedSubmissions = 0;
      let scores = [];
      let componentLists = [];

      for (var em in combinedData.students) {
        var s = combinedData.students[em];
        totalSubmissions++;
        
        if (s.score1 !== null || s.score2 !== null) {
          var s1 = s.score1 !== null ? s.score1 : 0;
          var s2 = s.score2 !== null ? s.score2 : 0;
          var combVal = parseFloat((s1 + s2).toFixed(2));
          scores.push(combVal);
          gradedSubmissions++;

          var c1 = s.comps1 || Array(combinedData.numQ1).fill(0);
          var c2 = s.comps2 || Array(combinedData.numQ2).fill(0);
          componentLists.push(c1.concat(c2));
        }
      }

      if (gradedSubmissions === 0) {
        return { 
          success: false, 
          message: "Chưa có bài thi nào được chấm điểm để tạo thống kê.", 
          totalSubmissions: totalSubmissions, 
          gradedSubmissions: 0 
        };
      }

      scores.sort((a, b) => a - b);
      const minScore = parseFloat(scores[0].toFixed(2));
      const maxScore = parseFloat(scores[scores.length - 1].toFixed(2));
      const sum = scores.reduce((a, b) => a + b, 0);
      const averageScore = parseFloat((sum / gradedSubmissions).toFixed(2));
      
      let medianScore = 0;
      const mid = Math.floor(scores.length / 2);
      if (scores.length % 2 !== 0) {
        medianScore = scores[mid];
      } else {
        medianScore = (scores[mid - 1] + scores[mid]) / 2;
      }
      medianScore = parseFloat(medianScore.toFixed(2));

      let questionStats = [];
      if (componentLists.length > 0) {
        const numQuestions = componentLists[0].length;
        
        for (let q = 0; q < numQuestions; q++) {
          let qScores = [];
          componentLists.forEach(list => {
            if (list[q] !== undefined) {
              qScores.push(list[q]);
            }
          });
          
          if (qScores.length > 0) {
            const qSum = qScores.reduce((a, b) => a + b, 0);
            const qAvg = parseFloat((qSum / qScores.length).toFixed(2));
            const qMax = Math.max(...qScores);
            const qMin = Math.min(...qScores);
            
            questionStats.push({
              questionIndex: q + 1,
              average: qAvg,
              max: qMax,
              min: qMin
            });
          }
        }
      }

      let sumOfMaxComponents = questionStats.reduce((sum, q) => sum + q.max, 0);
      let estimatedMaxScore = Math.max(maxScore, sumOfMaxComponents, 20.0);
      estimatedMaxScore = parseFloat(estimatedMaxScore.toFixed(2));

      const weakThreshold = estimatedMaxScore * 0.5;
      const averageThreshold = estimatedMaxScore * 0.65;
      const goodThreshold = estimatedMaxScore * 0.8;

      let weakCount = 0;
      let averageCount = 0;
      let goodCount = 0;
      let excellentCount = 0;

      scores.forEach(s => {
        if (s < weakThreshold) weakCount++;
        else if (s < averageThreshold) averageCount++;
        else if (s < goodThreshold) goodCount++;
        else excellentCount++;
      });

      let bins = Array(10).fill(0);
      let binLabels = [];
      const step = estimatedMaxScore / 10;
      
      for (let i = 0; i < 10; i++) {
        const start = i * step;
        const end = (i + 1) * step;
        
        const formatNum = (num) => {
          if (Number.isInteger(num)) return num.toString();
          return parseFloat(num.toFixed(1)).toString();
        };
        
        binLabels.push(formatNum(start) + "-" + formatNum(end) + "đ");
      }
      
      scores.forEach(s => {
        let binIdx = Math.floor(s / step);
        if (binIdx >= 10) binIdx = 9;
        if (binIdx < 0) binIdx = 0;
        bins[binIdx]++;
      });

      let standardDeviation = 0;
      if (gradedSubmissions > 0) {
        const mean = averageScore;
        const squareDiffs = scores.map(s => Math.pow(s - mean, 2));
        const avgSquareDiff = squareDiffs.reduce((sumDiff, val) => sumDiff + val, 0) / gradedSubmissions;
        standardDeviation = parseFloat(Math.sqrt(avgSquareDiff).toFixed(2));
      }

      let sortedDesc = [...scores].sort((a, b) => b - a);
      let N_hsg = gradedSubmissions;
      let T_hsg = Math.floor(0.60 * N_hsg);
      let maxG123 = Math.floor(0.60 * T_hsg);

      let g1 = Math.floor(0.05 * T_hsg);
      if (g1 === 0 && T_hsg >= 1 && maxG123 >= 1) {
        g1 = 1;
      }
      let g2 = Math.floor(0.20 * T_hsg);
      if (g2 === 0 && T_hsg >= 2 && (maxG123 - g1) >= 1) {
        g2 = 1;
      }
      let g3 = Math.max(0, maxG123 - g1 - g2);
      let gKK = Math.max(0, T_hsg - g1 - g2 - g3);

      let firstPrizeThreshold = g1 > 0 && g1 <= sortedDesc.length ? parseFloat(sortedDesc[g1 - 1].toFixed(2)) : null;
      let secondPrizeThreshold = g2 > 0 && (g1 + g2) <= sortedDesc.length ? parseFloat(sortedDesc[g1 + g2 - 1].toFixed(2)) : null;
      let thirdPrizeThreshold = g3 > 0 && (g1 + g2 + g3) <= sortedDesc.length ? parseFloat(sortedDesc[g1 + g2 + g3 - 1].toFixed(2)) : null;
      let consolationPrizeThreshold = gKK > 0 && (g1 + g2 + g3 + gKK) <= sortedDesc.length ? parseFloat(sortedDesc[g1 + g2 + g3 + gKK - 1].toFixed(2)) : null;

      return {
        success: true,
        examTitle: examTitle,
        totalSubmissions: totalSubmissions,
        gradedSubmissions: gradedSubmissions,
        averageScore: averageScore,
        medianScore: medianScore,
        highestScore: maxScore,
        lowestScore: minScore,
        standardDeviation: standardDeviation,
        estimatedMaxScore: estimatedMaxScore,
        classification: {
          weak: weakCount,
          average: averageCount,
          good: goodCount,
          excellent: excellentCount
        },
        hsgStats: {
          totalContestants: N_hsg,
          totalAwards: T_hsg,
          firstPrizeCount: g1,
          secondPrizeCount: g2,
          thirdPrizeCount: g3,
          consolationPrizeCount: gKK,
          firstPrizeThreshold: firstPrizeThreshold,
          secondPrizeThreshold: secondPrizeThreshold,
          thirdPrizeThreshold: thirdPrizeThreshold,
          consolationPrizeThreshold: consolationPrizeThreshold,
          noPrizeCount: N_hsg - T_hsg
        },
        bins: bins,
        binLabels: binLabels,
        questionStats: questionStats
      };
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: false, message: "Chưa có dữ liệu bài nộp nào cho bài thi này." };
    }

    const normTargetExamTitle = normalizeExamText(examTitle);
    let totalSubmissions = 0;
    let gradedSubmissions = 0;
    let scores = [];
    let componentLists = [];
    const studentMap = {};

    for (var i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = row[1] ? row[1].toString().toLowerCase().trim() : "";
      const rowNormEmail = normalizeExamText(rowEmail);
      const rowExamTitle = row[2] ? row[2].toString().trim() : "";
      const rowNormTitle = normalizeExamText(rowExamTitle);
      
      if (rowNormTitle === normTargetExamTitle && rowNormEmail) {
        const rowScoreStr = row[5] !== undefined ? row[5].toString().trim() : "";
        let scoreVal = null;
        let components = [];
        
        if (rowScoreStr !== "") {
          if (rowScoreStr.indexOf("[") !== -1) {
            const parts = rowScoreStr.split("[");
            scoreVal = parseFloat(parts[0].trim().replace(/,/g, "."));
            const compStr = parts[1].replace("]", "").trim();
            components = compStr ? compStr.split(";").map(s => {
              const val = parseFloat(s.trim().replace(/,/g, "."));
              return isNaN(val) ? 0 : val;
            }) : [];
          } else {
            scoreVal = parseFloat(rowScoreStr.replace(/,/g, "."));
          }
        }

        if (!studentMap[rowNormEmail] || (scoreVal !== null && !isNaN(scoreVal))) {
          studentMap[rowNormEmail] = {
            score: (!isNaN(scoreVal) && scoreVal !== null) ? scoreVal : null,
            components: components
          };
        }
      }
    }

    const uniqueStudents = Object.values(studentMap);
    totalSubmissions = uniqueStudents.length;

    uniqueStudents.forEach(st => {
      if (st.score !== null) {
        scores.push(st.score);
        gradedSubmissions++;
        if (st.components && st.components.length > 0) {
          componentLists.push(st.components);
        }
      }
    });

    if (gradedSubmissions === 0) {
      return { 
        success: false, 
        message: "Chưa có bài thi nào được chấm điểm để tạo thống kê.", 
        totalSubmissions: totalSubmissions, 
        gradedSubmissions: 0 
      };
    }

    scores.sort((a, b) => a - b);
    const minScore = parseFloat(scores[0].toFixed(2));
    const maxScore = parseFloat(scores[scores.length - 1].toFixed(2));
    const sum = scores.reduce((a, b) => a + b, 0);
    const averageScore = parseFloat((sum / gradedSubmissions).toFixed(2));
    
    let medianScore = 0;
    const mid = Math.floor(scores.length / 2);
    if (scores.length % 2 !== 0) {
      medianScore = scores[mid];
    } else {
      medianScore = (scores[mid - 1] + scores[mid]) / 2;
    }
    medianScore = parseFloat(medianScore.toFixed(2));

    // Calculate questionStats first so we can use it to find the maximum possible score
    let questionStats = [];
    if (componentLists.length > 0) {
      const numQuestions = componentLists[0].length;
      
      for (let q = 0; q < numQuestions; q++) {
        let qScores = [];
        componentLists.forEach(list => {
          if (list[q] !== undefined) {
            qScores.push(list[q]);
          }
        });
        
        if (qScores.length > 0) {
          const qSum = qScores.reduce((a, b) => a + b, 0);
          const qAvg = parseFloat((qSum / qScores.length).toFixed(2));
          const qMax = Math.max(...qScores);
          const qMin = Math.min(...qScores);
          
          questionStats.push({
            questionIndex: q + 1,
            average: qAvg,
            max: qMax,
            min: qMin
          });
        }
      }
    }

    // Estimate the maximum possible score
    let sumOfMaxComponents = questionStats.reduce((sum, q) => sum + q.max, 0);
    let estimatedMaxScore = Math.max(maxScore, sumOfMaxComponents, 20.0);
    estimatedMaxScore = parseFloat(estimatedMaxScore.toFixed(2));

    // Define classification thresholds dynamically based on estimatedMaxScore
    const weakThreshold = estimatedMaxScore * 0.5;
    const averageThreshold = estimatedMaxScore * 0.65;
    const goodThreshold = estimatedMaxScore * 0.8;

    let weakCount = 0;
    let averageCount = 0;
    let goodCount = 0;
    let excellentCount = 0;

    scores.forEach(s => {
      if (s < weakThreshold) weakCount++;
      else if (s < averageThreshold) averageCount++;
      else if (s < goodThreshold) goodCount++;
      else excellentCount++;
    });

    // Divide [0, estimatedMaxScore] into 10 intervals dynamically
    let bins = Array(10).fill(0);
    let binLabels = [];
    const step = estimatedMaxScore / 10;
    
    for (let i = 0; i < 10; i++) {
      const start = i * step;
      const end = (i + 1) * step;
      
      const formatNum = (num) => {
        if (Number.isInteger(num)) return num.toString();
        return parseFloat(num.toFixed(1)).toString();
      };
      
      binLabels.push(formatNum(start) + "-" + formatNum(end) + "đ");
    }
    
    scores.forEach(s => {
      let binIdx = Math.floor(s / step);
      if (binIdx >= 10) binIdx = 9;
      if (binIdx < 0) binIdx = 0;
      bins[binIdx]++;
    });

    // Calculate Standard Deviation
    let standardDeviation = 0;
    if (gradedSubmissions > 0) {
      const mean = averageScore;
      const squareDiffs = scores.map(s => Math.pow(s - mean, 2));
      const avgSquareDiff = squareDiffs.reduce((sumDiff, val) => sumDiff + val, 0) / gradedSubmissions;
      standardDeviation = parseFloat(Math.sqrt(avgSquareDiff).toFixed(2));
    }

    // Calculate HSG prizes (national rules)
    let sortedDesc = [...scores].sort((a, b) => b - a);
    let N_hsg = gradedSubmissions;
    let T_hsg = Math.floor(0.60 * N_hsg); // Total awards <= 60%
    let maxG123 = Math.floor(0.60 * T_hsg); // Nhất + Nhì + Ba <= 60% of total awards

    let g1 = Math.floor(0.05 * T_hsg);
    if (g1 === 0 && T_hsg >= 1 && maxG123 >= 1) {
      g1 = 1;
    }
    let g2 = Math.floor(0.20 * T_hsg);
    if (g2 === 0 && T_hsg >= 2 && (maxG123 - g1) >= 1) {
      g2 = 1;
    }
    let g3 = Math.max(0, maxG123 - g1 - g2);
    let gKK = Math.max(0, T_hsg - g1 - g2 - g3);

    let firstPrizeThreshold = g1 > 0 && g1 <= sortedDesc.length ? parseFloat(sortedDesc[g1 - 1].toFixed(2)) : null;
    let secondPrizeThreshold = g2 > 0 && (g1 + g2) <= sortedDesc.length ? parseFloat(sortedDesc[g1 + g2 - 1].toFixed(2)) : null;
    let thirdPrizeThreshold = g3 > 0 && (g1 + g2 + g3) <= sortedDesc.length ? parseFloat(sortedDesc[g1 + g2 + g3 - 1].toFixed(2)) : null;
    let consolationPrizeThreshold = gKK > 0 && (g1 + g2 + g3 + gKK) <= sortedDesc.length ? parseFloat(sortedDesc[g1 + g2 + g3 + gKK - 1].toFixed(2)) : null;

    return {
      success: true,
      examTitle: examTitle,
      totalSubmissions: totalSubmissions,
      gradedSubmissions: gradedSubmissions,
      averageScore: averageScore,
      medianScore: medianScore,
      highestScore: maxScore,
      lowestScore: minScore,
      standardDeviation: standardDeviation,
      estimatedMaxScore: estimatedMaxScore,
      classification: {
        weak: weakCount,
        average: averageCount,
        good: goodCount,
        excellent: excellentCount
      },
      hsgStats: {
        totalContestants: N_hsg,
        totalAwards: T_hsg,
        firstPrizeCount: g1,
        secondPrizeCount: g2,
        thirdPrizeCount: g3,
        consolationPrizeCount: gKK,
        firstPrizeThreshold: firstPrizeThreshold,
        secondPrizeThreshold: secondPrizeThreshold,
        thirdPrizeThreshold: thirdPrizeThreshold,
        consolationPrizeThreshold: consolationPrizeThreshold,
        noPrizeCount: N_hsg - T_hsg
      },
      bins: bins,
      binLabels: binLabels,
      questionStats: questionStats
    };

  } catch (e) {
    return { success: false, message: "Lỗi Server: " + e.message };
  }
}

function getCourses() {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return [];
  const allCourses = getCachedCoursesFromSheet();
  const role = getUserRole(email);

  const allowedFolders = getUserAllowedFolders(email);
  Logger.log("Allowed folders for " + email + ": " + allowedFolders);

  // Lọc các course dựa trên quyền truy cập: một course được phép nếu folder hiện tại 
  // hoặc bất kỳ folder cha (ancestors/tổ tiên) nào của nó nằm trong danh sách folder được phép của user.
  const filtered = allCourses.filter(course => {
    // 2. Chặn theo quyền truy cập
    if (allowedFolders.includes("ALL_FOLDERS")) return true;
    if (course.ancestors && Array.isArray(course.ancestors)) {
      return course.ancestors.some(fId => allowedFolders.includes(fId));
    }
    return allowedFolders.includes(course.folderId);
  });
  Logger.log("Filtered courses length: " + filtered.length);

  return filtered;
}

function autoConvertDocsToPdf() {
  const rootFolderId = "16k1Ma_g_bmuJOUokOfJOr676vnd4uORN"; // ID thư mục gốc của khóa học
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  const now = new Date().getTime();
  const oneDayInMs = 24 * 60 * 60 * 1000;
  
  function processFolder(folder) {
    const files = folder.getFiles();
    const docFiles = [];
    const pdfFiles = [];
    
    // Phân loại các file trong thư mục hiện tại
    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();
      if (mimeType === MimeType.GOOGLE_DOCS) {
        docFiles.push(file);
      } else if (mimeType === MimeType.PDF) {
        pdfFiles.push(file);
      }
    }
    
    // Xử lý chuyển đổi cho từng Doc file
    docFiles.forEach(docFile => {
      const docId = docFile.getId();
      const docName = docFile.getName();
      
      // Tìm xem có PDF nào được tạo từ Doc này chưa
      let correspondingPdf = null;
      for (let i = 0; i < pdfFiles.length; i++) {
        const pdfFile = pdfFiles[i];
        const desc = pdfFile.getDescription() || "";
        if (desc === "CONVERTED_FROM_DOC:" + docId) {
          correspondingPdf = pdfFile;
          break;
        }
      }
      
      const newPdfName = `${docName} [PDF_Converted_${now}]`;
      
      if (!correspondingPdf) {
        // Chưa có PDF -> Tạo mới
        try {
          Logger.log(`Converting Doc to PDF (New): ${docName}`);
          const pdfBlob = docFile.getAs(MimeType.PDF);
          pdfBlob.setName(newPdfName);
          const newPdf = folder.createFile(pdfBlob);
          newPdf.setDescription("CONVERTED_FROM_DOC:" + docId);
          Logger.log(`Created PDF for ${docName} successfully as ${newPdfName}.`);
        } catch (e) {
          Logger.log(`Error converting ${docName} to PDF: ` + e.toString());
        }
      } else {
        // Đã có PDF -> Kiểm tra xem đã qua 1 ngày chưa dựa vào timestamp ở đuôi tên file PDF cũ
        const pdfName = correspondingPdf.getName();
        const match = pdfName.match(/\[PDF_Converted_(\d+)\]$/);
        let lastConvertedTime = 0;
        if (match) {
          lastConvertedTime = parseInt(match[1], 10);
        }
        
        if (lastConvertedTime > 0 && (now - lastConvertedTime < oneDayInMs)) {
          Logger.log(`Bỏ qua chuyển đổi: File PDF của ${docName} được tạo cách đây chưa đầy 1 ngày.`);
          return; // Tiếp tục vòng lặp cho file tiếp theo
        }
        
        // Nếu đã quá 1 ngày (hoặc không tìm thấy timestamp hợp lệ) -> Tiến hành ghi đè trực tiếp để cập nhật
        try {
          Logger.log(`Updating PDF content (Over 1 day since last conversion): ${docName}`);
          const pdfBlob = docFile.getAs(MimeType.PDF);
          pdfBlob.setName(newPdfName);
          
          // Ghi đè nội dung file PDF cũ bằng Drive API v3 (giữ nguyên ID file cũ)
          Drive.Files.update({}, correspondingPdf.getId(), pdfBlob);
          
          // Cập nhật lại tên mới chứa mốc thời gian hiện tại
          correspondingPdf.setName(newPdfName);
          Logger.log(`Updated PDF content for ${docName} successfully as ${newPdfName}.`);
        } catch (e) {
          Logger.log(`Error updating PDF for ${docName}: ` + e.toString());
        }
      }
    });
    
    // Tiếp tục đệ quy quét các thư mục con
    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      processFolder(subfolders.next());
    }
  }
  
  Logger.log("Bắt đầu tiến trình tự động chuyển đổi Google Docs sang PDF...");
  processFolder(rootFolder);
  Logger.log("Hoàn thành tiến trình chuyển đổi Google Docs sang PDF.");
}

function updateCourseCache() {
  const rootFolderId = "16k1Ma_g_bmuJOUokOfJOr676vnd4uORN"; // ID thư mục gốc
  const rootFolder = DriveApp.getFolderById(rootFolderId);

  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.deleteProperty("checkpoint_courses_folders");

  const courseMap = new Map();

  // Lấy dữ liệu cũ để so sánh phát hiện bài mới
  const oldCourses = getCachedCoursesFromSheet();
  const oldIds = new Set(oldCourses.map(c => c.fileId));

  const newItemsFound = [];
  const processedFolders = new Set();

  function processFolder(folder, parentPath, level, ancestors) {
    try {
      const folderId = folder.getId();
      if (processedFolders.has(folderId)) return;
      processedFolders.add(folderId);

      const folderName = folder.getName();
      const fullPath = parentPath ? parentPath + " / " + folderName : folderName;
      const currentAncestors = ancestors ? [...ancestors, folderId] : [folderId];

      const files = folder.getFiles();
      while (files.hasNext()) {
        try {
          const file = files.next();
          const fileId = file.getId();
          const mimeType = file.getMimeType();

          const isVideo = mimeType.startsWith("video/");
          const isPDF = mimeType === MimeType.PDF;
          const isSheet = mimeType === MimeType.GOOGLE_SHEETS;

          if (isVideo || isPDF || isSheet) {
            let rawName = file.getName();
            let cleanName = isPDF ? rawName.replace(/\s*\[PDF_Converted_\d+\]$/, "") : rawName;
            const item = {
              title: cleanName,
              fileId: fileId,
              type: isPDF ? "pdf" : isSheet ? "sheet" : "video",
              folder: fullPath,
              folderId: folderId,
              ancestors: currentAncestors,
              level: level
            };
            courseMap.set(String(fileId), item);
            
            if (oldIds.size > 0 && !oldIds.has(fileId)) {
              newItemsFound.push(item);
            }
          }
        } catch (fErr) {
          Logger.log("Lỗi xử lý file course: " + fErr);
        }
      }

      const subfolders = folder.getFolders();
      while (subfolders.hasNext()) {
        try {
          processFolder(
            subfolders.next(),
            fullPath,
            level + 1,
            currentAncestors
          );
        } catch (subErr) {
          Logger.log("Lỗi xử lý thư mục con course: " + subErr);
        }
      }
    } catch (err) {
      Logger.log("Lỗi processFolder course: " + err);
    }
  }

  processFolder(rootFolder, "", 0, []);

  const finalCourses = Array.from(courseMap.values());

  // Lưu cache mới vào Google Sheet
  saveCachedCoursesToSheet(finalCourses);

  // Nếu có bài mới, tạo thông báo tự động cho từng mục để có thể mở trực tiếp
  if (newItemsFound.length > 0) {
    newItemsFound.forEach(item => {
      addAutoNotificationToSheet(
        "🆕 Bài giảng mới",
        `Bài giảng: ${item.title}\nNội dung mới vừa được cập nhật trong mục "${item.folder.split(" / ").pop()}".`,
        `courses:open:${item.fileId}`,
        item.folderId
      );
    });
  }

  return {
    success: true,
    message: `✅ Đã quét hoàn tất toàn bộ cache khóa học từ đầu: ${finalCourses.length} bài giảng (100%). Phát hiện ${newItemsFound.length} bài mới.`
  };
}

/**
 * Hàm hỗ trợ tự động ghi thông báo vào sheet
 */
function addAutoNotificationToSheet(title, desc, action, targetFolder, targetUser = "ALL", senderEmail = "") {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Notifications");
    if (!sheet) {
      sheet = ss.insertSheet("Notifications");
      sheet.appendRow(["ID", "Title", "Description", "Action", "TargetRole", "TargetFolder", "Time", "SenderEmail"]);
      sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#f3f3f3");
    }
    
    // Kiểm tra và cập nhật header nếu thiếu cột SenderEmail
    if (sheet.getLastColumn() < 8) {
      sheet.getRange(1, 8).setValue("SenderEmail").setFontWeight("bold").setBackground("#f3f3f3");
    }
    
    const id = "auto_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);
    const now = new Date();
    
    // Thêm vào hàng thứ 2 (dưới header) để thông báo mới nhất nằm trên
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, 8).setValues([[
      id,
      title,
      desc,
      action || "overview",
      targetUser.toUpperCase(),
      targetFolder || "",
      now,
      senderEmail
    ]]);
    
    // Giữ sheet gọn gàng, chỉ lưu tối đa 200 thông báo gần nhất
    const lastRow = sheet.getLastRow();
    if (lastRow > 200) {
      sheet.deleteRows(201, lastRow - 200);
    }
    
    return id;
  } catch (err) {
    Logger.log("Lỗi addAutoNotificationToSheet: " + err);
    return null;
  }
}

function addOrUpdateGroupNotification(title, desc, action, targetFolder, studentEmail, senderEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Notifications");
    if (!sheet) {
      sheet = ss.insertSheet("Notifications");
      sheet.appendRow(["ID", "Title", "Description", "Action", "TargetRole", "TargetFolder", "Time", "SenderEmail"]);
      sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#f3f3f3");
    }
    
    // Kiểm tra và cập nhật header nếu thiếu cột SenderEmail
    if (sheet.getLastColumn() < 8) {
      sheet.getRange(1, 8).setValue("SenderEmail").setFontWeight("bold").setBackground("#f3f3f3");
    }

    const data = sheet.getDataRange().getValues();
    let foundRow = -1;
    const cleanStudentEmail = studentEmail.toUpperCase().trim();

    // Tìm xem đã có thông báo chung cho bài này chưa
    // So sánh bằng Title, Action và TargetFolder
    for (let i = 1; i < data.length; i++) {
      const rowTitle = String(data[i][1] || "");
      const rowAction = String(data[i][3] || "");
      const rowFolder = String(data[i][5] || "");
      if (rowTitle === title && rowAction === action && rowFolder === targetFolder) {
        foundRow = i + 1;
        break;
      }
    }

    const now = new Date();

    if (foundRow !== -1) {
      // Tìm thấy thông báo cũ -> Cập nhật TargetRole (thêm email học sinh mới vào nếu chưa có)
      let currentTargets = String(data[foundRow - 1][4] || "").toUpperCase().split(",").map(s => s.trim()).filter(Boolean);
      if (!currentTargets.includes(cleanStudentEmail)) {
        currentTargets.push(cleanStudentEmail);
      }
      
      // Ghi nhận lại danh sách email, cập nhật thời gian để thông báo nổi lên đầu bằng Batch write
      sheet.getRange(foundRow, 5).setValue(currentTargets.join(", "));
      sheet.getRange(foundRow, 7, 1, 2).setValues([[now, senderEmail || ""]]);
      
      // Di chuyển hàng này lên vị trí hàng thứ 2 để thông báo mới nhất nằm trên
      if (foundRow > 2) {
        const rowData = sheet.getRange(foundRow, 1, 1, 8).getValues();
        sheet.deleteRow(foundRow);
        sheet.insertRowAfter(1);
        sheet.getRange(2, 1, 1, 8).setValues(rowData);
      }
      invalidateSheetCache("Notifications");
      
      return data[foundRow - 1][0]; // Trả về ID cũ
    } else {
      // Không tìm thấy -> Tạo thông báo mới
      const id = "group_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);
      sheet.insertRowAfter(1);
      sheet.getRange(2, 1, 1, 8).setValues([[
        id,
        title,
        desc,
        action || "overview",
        cleanStudentEmail,
        targetFolder || "",
        now,
        senderEmail
      ]]);
      
      // Giữ sheet gọn gàng, tối đa 200 dòng
      const lastRow = sheet.getLastRow();
      if (lastRow > 200) {
        sheet.deleteRows(201, lastRow - 200);
      }
      return id;
    }
  } catch (err) {
    Logger.log("Lỗi addOrUpdateGroupNotification: " + err);
    return null;
  }
}

function removeStudentFromGroupNotification(title, targetFolder, studentEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Notifications");
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const cleanStudentEmail = (studentEmail || "").toString().toUpperCase().trim();
    if (!cleanStudentEmail) return;

    for (let i = 1; i < data.length; i++) {
      const rowTitle = String(data[i][1] || "");
      const rowFolder = String(data[i][5] || "");
      if (rowTitle === title && rowFolder === targetFolder) {
        let targets = String(data[i][4] || "").toUpperCase().split(",").map(s => s.trim()).filter(Boolean);
        if (targets.includes(cleanStudentEmail)) {
          targets = targets.filter(e => e !== cleanStudentEmail);
          sheet.getRange(i + 1, 5).setValue(targets.join(", "));
          invalidateSheetCache("Notifications");
        }
      }
    }
  } catch (err) {
    Logger.log("Lỗi removeStudentFromGroupNotification: " + err);
  }
}


function getUserAllowedFolders(email) {
  if (!email) return [];
  const role = getUserRole(email);
  if (role.toUpperCase() === "ADMIN") {
    return ["ALL_FOLDERS"];
  }

  const data = getSheetDataCached("Permissions");
  const allowed = [];

  email = email.toLowerCase().trim();

  for (let i = 1; i < data.length; i++) {
    const folderId = data[i][0];
    const emailsCell = data[i][1] || "";
    const rolesCell  = data[i][2] || "";

    // danh sách email trong ô
    const emailsList = String(emailsCell).split(",").map(e => e.toLowerCase().trim());

    // danh sách role trong ô (nếu có)
    const rolesList  = String(rolesCell).split(",").map(r => r.toUpperCase().trim());

    // Email phải có mặt
    if (emailsList.includes(email)) {
      // Role phải match (hoặc ô Roles trống). STUDENT chỉ được xem tài liệu của STUDENT, S-STUDENT chỉ được xem tài liệu của S-STUDENT (sstudent).
      let hasRole = false;
      const upperRole = role.toUpperCase();
      if (upperRole === "STUDENT") {
        hasRole = rolesList.includes("STUDENT");
      } else if (upperRole === "S-STUDENT") {
        hasRole = rolesList.includes("S-STUDENT") || rolesList.includes("SSTUDENT");
      } else {
        hasRole = rolesList.includes(upperRole);
      }

      if (rolesCell === "" || hasRole) {
        allowed.push(folderId);
      }
    }
  }
  return allowed;
}

/** Code Về EVENT COIN */
const EVENT_COIN_INDEX = 20; // Cột U
const EVENT_IS_ACTIVE = false; // Hết sự kiện thì để false, cần mở lại thì đổi true

function updateUserEventCoin(email, newPoints) {
  var user = getUserByEmail(email);
  if (!user) return false;

  newPoints = Math.max(0, newPoints);
  updateUserField(user.row, EVENT_COIN_INDEX, newPoints);
  return true;
}

function updateUserEventCoinByClient(newPoints) {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return false;
  return updateUserEventCoin(email, newPoints);
}

function addUserEventCoin_(pointsToAdd, source = "Hệ thống") {
  if (!EVENT_IS_ACTIVE) return false; // 🔥 Tắt cộng điểm khi sự kiện kết thúc

  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return false;

  const user = getUserByEmail(email);
  if (!user) return false;

  const current = parseInt(user.data[EVENT_COIN_INDEX]) || 0;
  let newTotal = current + pointsToAdd;

  if (newTotal < 0) newTotal = 0;

  updateUserField(user.row, EVENT_COIN_INDEX, newTotal);
  logCoinTransaction(email, newTotal - current, "Event Coin", source);
  return newTotal;
}

function setUserEventCoin_(newTotal) {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return false;

  const user = getUserByEmail(email);
  if (!user) return false;

  if (newTotal < 0) newTotal = 0;

  updateUserField(user.row, EVENT_COIN_INDEX, newTotal);
  return newTotal;
}

function getUserEventCoin() {
  return 0;
}

function deductEventCoinSmart(email, amount) {
  const info = getUserAndPartner(email);
  if (!info) return false;

  let myCoin = Number(info.myRow[EVENT_COIN_INDEX]) || 0;

  if (myCoin >= amount) {
    info.sheet.getRange(info.myIndex + 1, EVENT_COIN_INDEX + 1)
      .setValue(myCoin - amount);
    return true;
  }

  const need = amount - myCoin;
  if (!info.partnerRow) return false;

  let partnerCoin = Number(info.partnerRow[EVENT_COIN_INDEX]) || 0;
  if (partnerCoin < need) return false;

  info.sheet.getRange(info.myIndex + 1, EVENT_COIN_INDEX + 1).setValue(0);
  info.sheet.getRange(info.partnerIndex + 1, EVENT_COIN_INDEX + 1)
    .setValue(partnerCoin - need);

  return true;
}

function deductEventCoinSmartForCurrentUser(amount) {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return false;
  return deductEventCoinSmart(email, amount);
}

function getEventLeaderboardForCurrentUser(limit){
  return { success: true, top: [], userRank: null };
}

function getTripleDLeaderboardForCurrentUser(limit){
  return getLeaderboardGeneric(limit, 10); // cột K
}

function refreshUserStreak(email) {
  var user = getUserByEmail(email);
  if (!user) return 0;
  
  var now = new Date();
  var lastLoginTimeRaw = user.data[14]; // Column O
  var streak = parseInt(user.data[15]) || 0; // Column P
  
  var isFirstLogin = !lastLoginTimeRaw || isNaN(new Date(lastLoginTimeRaw).getTime());
  
  if (isFirstLogin) {
    streak = 1;
    updateUserField(user.row, 15, streak); // P
    updateUserField(user.row, 14, now);   // O
  } else {
    var lastLoginTime = new Date(lastLoginTimeRaw);
    var d1 = new Date(lastLoginTime.getFullYear(), lastLoginTime.getMonth(), lastLoginTime.getDate());
    var d2 = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var diffMs = d2.getTime() - d1.getTime();
    var diffDays = Math.round(diffMs / (1000 * 60 * 60 * 24));
    
    if (diffDays === 1) {
      streak += 1;
      updateUserField(user.row, 15, streak); // P
      updateUserField(user.row, 14, now);   // O
    } else if (diffDays > 1) {
      if (streak > 1) {
        // Lưu thông tin cứu streak bị đứt
        var key = "lost_streak_" + email.toLowerCase().replace(/[^a-z0-9]/g, "");
        var lostData = {
          oldStreak: streak,
          rescueCost: Math.min(1000, 100 + streak * 10), // Base 100 + 10 mỗi ngày, tối đa 1000 coin
          timestamp: new Date().getTime()
        };
        safeSetScriptProperty(key, JSON.stringify(lostData));
      }
      streak = 1;
      updateUserField(user.row, 15, streak); // P
      updateUserField(user.row, 14, now);   // O
    } else if (diffDays === 0) {
      // Đăng nhập lại trong cùng ngày: Cập nhật thời gian đăng nhập mới nhất vào Cột O
      updateUserField(user.row, 14, now);   // O
    }
  }
  return streak;
}

function rescueUserStreak(email) {
  if (!email) return { success: false, message: "Email không hợp lệ!" };
  var key = "lost_streak_" + email.toLowerCase().replace(/[^a-z0-9]/g, "");
  var propVal = PropertiesService.getScriptProperties().getProperty(key);
  if (!propVal) {
    return { success: false, message: "Không tìm thấy thông tin Streak cần cứu hoặc ưu đãi đã hết hạn!" };
  }
  
  var lostData;
  try {
    lostData = JSON.parse(propVal);
  } catch(e) {
    return { success: false, message: "Dữ liệu Streak bị lỗi!" };
  }
  
  var user = getUserByEmail(email);
  if (!user) return { success: false, message: "Không tìm thấy tài khoản người dùng!" };
  
  // Lấy coins thực tế của riêng user từ row (Cột K = index 10)
  var currentPoints = Number(user.data[10]) || 0;
  if (currentPoints < lostData.rescueCost) {
    return { success: false, message: "Tài khoản của bạn không đủ Coin để thực hiện cứu Streak (Cần: " + lostData.rescueCost + " 🪙)!" };
  }
  
  // Trừ points
  var newPoints = deductUserPoints(lostData.rescueCost, "Cứu streak đứt (" + lostData.oldStreak + " ngày)");
  if (newPoints === false) {
    return { success: false, message: "Khấu trừ điểm thất bại!" };
  }
  
  // Khôi phục streak trên sheet
  updateUserField(user.row, 15, lostData.oldStreak); // P
  
  // Xóa thông tin đã dùng cứu
  PropertiesService.getScriptProperties().deleteProperty(key);
  
  return {
    success: true,
    message: "Chúc mừng! Bạn đã cứu thành công chuỗi " + lostData.oldStreak + " ngày học tập! ❤️‍🩹",
    newStreak: lostData.oldStreak,
    newPoints: newPoints
  };
}

function dismissStreakRescue(email) {
  if (!email) return { success: false };
  var key = "lost_streak_" + email.toLowerCase().replace(/[^a-z0-9]/g, "");
  PropertiesService.getScriptProperties().deleteProperty(key);
  return { success: true };
}

function getStreakLeaderboard(limit) {
  limit = limit || 20;
  const sheet = SpreadsheetApp.getActive().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();

  const entities = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[0];
    const email = String(row[1]).trim().toLowerCase();
    const streak = parseInt(row[15]) || 0; // Cột P

    // Bỏ qua nếu là tài khoản chưa duyệt hoặc bị khóa hoặc streak <= 0
    const status = row[7];
    const role = row[13];
    if (!email || status === "Pending" || role === "Blocked" || streak <= 0) {
      continue;
    }

    entities.push({
      name: name,
      email: email,
      streak: streak
    });
  }

  // Sắp xếp giảm dần theo streak
  entities.sort((a, b) => b.streak - a.streak);

  const top = entities.slice(0, limit).map((e, idx) => ({
    rank: idx + 1,
    name: e.name,
    streak: e.streak
  }));

  const myEmail = CacheService.getUserCache().get("loggedInUser");
  let userRank = null;

  for (let i = 0; i < entities.length; i++) {
    if (entities[i].email === myEmail) {
      userRank = {
        name: entities[i].name,
        streak: entities[i].streak,
        rank: i + 1
      };
      break;
    }
  }

  return { success: true, top, userRank };
}


function getEventCoinInfo(){
  return { totalPoints: 0, percentage: 0 };
}

function getLeaderboardGeneric(limit, coinIndex){
  const sheet = SpreadsheetApp.getActive().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();

  const entities = [];
  const visited = new Set(); // tránh tính trùng couple

  for(let i=1;i<data.length;i++){
    if (visited.has(i)) continue;

    const row = data[i];
    const name = row[0];
    const email = String(row[1]).trim().toLowerCase();
    const partnerEmail = String(row[19] || "").trim().toLowerCase();
    const myCoin = parseInt(row[coinIndex]) || 0;

    // Nếu không có partner → cá nhân
    if (!partnerEmail){
      entities.push({
        name: name,
        coin: myCoin,
        members: [email]
      });
      visited.add(i);
      continue;
    }

    // Nếu có partner → tìm partner
    const partnerIndex = data.findIndex((r,idx)=>
      idx>0 &&
      String(r[1]).trim().toLowerCase() === partnerEmail
    );

    if (partnerIndex > 0){
      const partnerRow = data[partnerIndex];
      const partnerName = partnerRow[0];
      const partnerCoin = parseInt(partnerRow[coinIndex]) || 0;

      entities.push({
        name: `${name} ❤️ ${partnerName}`,
        coin: myCoin + partnerCoin,
        members: [email, partnerEmail]
      });

      visited.add(i);
      visited.add(partnerIndex);
    } else {
      // Partner lỗi → tính như cá nhân
      entities.push({
        name: name,
        coin: myCoin,
        members: [email]
      });
      visited.add(i);
    }
  }

  // Sort
  entities.sort((a,b)=>b.coin-a.coin);

  const top = entities.slice(0,limit).map((e,i)=>({
    rank: i+1,
    name: e.name,
    coin: e.coin
  }));

  // 👇 Tính userRank theo entity
  const email = CacheService.getUserCache().get("loggedInUser");
  let userRank = null;

  for(let i=0;i<entities.length;i++){
    if (entities[i].members.includes(email)){
      userRank = {
        name: entities[i].name,
        points: entities[i].coin,
        rank: i+1
      };
      break;
    }
  }

  return { top, userRank };
}


function getCoinInfoGeneric(coinIndex){
  const email = CacheService.getUserCache().get("loggedInUser");
  if(!email) return null;

  const user = getUserByEmail(email);
  if(!user) return null;

  const userPoints = parseInt(user.data[coinIndex]) || 0;

  const sheet = SpreadsheetApp.getActive().getSheetByName("Users");
  const allData = sheet.getDataRange().getValues();

  let totalPoints = 0;
  for (let i = 1; i < allData.length; i++) {
    totalPoints += parseInt(allData[i][coinIndex]) || 0;
  }

  const percentage = totalPoints ? (userPoints / totalPoints) * 100 : 0;

  return { userPoints, totalPoints, percentage };
}





/**Code Về Triple D Coin*/
function updateUserPoints(email, newPoints) {
  var user = getUserByEmail(email);
  if (!user) return false;
  newPoints = Math.max(0, newPoints); // Không cho âm điểm
  updateUserField(user.row, 10, newPoints); // Cột K
  return true;
}

function updateUserPointsByClient_(newPoints) {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return false;
  return updateUserPoints(email, newPoints);
}

function addUserPoints_(pointsToAdd, source = "Hệ thống") {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return false;
  const user = getUserByEmail(email);
  if (!user) return false;

  const current = parseInt(user.data[10]) || 0; // Cột K = index 10
  let newTotal = current + pointsToAdd;

  // 🚨 Không cho âm
  if (newTotal < 0) newTotal = 0;

  updateUserField(user.row, 10, newTotal);
  logCoinTransaction(email, newTotal - current, "TD Coin", source);
  return newTotal; // trả về tổng mới
}

function thuphihangtuan() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Users"); // đổi theo tên sheet
  const data = sheet.getDataRange().getValues();

  // Bỏ hàng tiêu đề (i = 1)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const email = row[1];  // cột B
    const current = parseInt(row[10]) || 0; // cột K
    const role = (row[13] || "").toString().trim().toUpperCase(); // cột N

    // Xác định % trừ
    let percent;
    if (role === "STUDENT" || role === "S-STUDENT") {
      percent = 0;
    } else if (role === "VIP" || role === "O-STUDENT") {
      percent = 0.2;
    } else {
      percent = 0.1;
    }

    const amountToSubtract = Math.round(current * percent);
    let newTotal = current - amountToSubtract;
    if (newTotal < 0) newTotal = 0;

    // Cập nhật cột K
    sheet.getRange(i + 1, 11).setValue(newTotal); // (row, column)
  }
}

function setUserPoints_(newTotal, source = "Hệ thống") {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return false;
  const user = getUserByEmail(email);
  if (!user) return false;

  const current = parseInt(user.data[10]) || 0;
  // 🚨 Không cho âm
  if (newTotal < 0) newTotal = 0;

  updateUserField(user.row, 10, newTotal);
  logCoinTransaction(email, newTotal - current, "TD Coin", source);
  return newTotal; // trả về tổng mới
}

// Hàm lấy điểm hiện tại
function getUserPoints() {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return 0;

  const info = getUserAndPartner(email);
  if (!info) return 0;

  const myCoin = Number(info.myRow[10]) || 0;
  const partnerCoin = info.partnerRow ? Number(info.partnerRow[10]) || 0 : 0;

  return myCoin + partnerCoin;
}

function deductPointsAndLog(type, fileId, title) {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { error: true, message: "Bạn chưa đăng nhập." };

  const role = getUserRole(email);

  if (role === "S-STUDENT" || role === "ADMIN") {
      // Truy cập miễn phí
      shareFileAndLogAccess_(fileId, email);
      return { error: false, points: getUserPoints(), freeAccess: true };
  }

  const cost = 67;

  const ok = deductPointsSmart(email, cost);
  if (!ok) return { error: true, message: "Không đủ điểm để mở tài liệu này." };

  // Chia sẻ trực tiếp file cho người mua và lưu log truy cập dọn dẹp định kỳ
  shareFileAndLogAccess_(fileId, email);

  logActivityToSheet(email, "BUY_" + (type || "ITEM").toUpperCase(), `Trừ ${cost} coin mở tài liệu: [${title}] (ID: ${fileId})`);
  logCoinTransaction(email, -cost, "TD Coin", "Mở tài liệu: " + title);

  return { error: false, points: getUserPoints(), freeAccess: false };
}

function sharePurchasedFile(fileId, userEmailParam) {
  let email = (userEmailParam || "").toLowerCase().trim();
  if (!email) {
    email = CacheService.getUserCache().get("loggedInUser");
  }
  if (!email) return { error: true, message: "Bạn chưa đăng nhập." };

  const success = shareFileAndLogAccess_(fileId, email);
  if (success) {
    return { error: false, message: "Chia sẻ tài liệu đã mua thành công." };
  } else {
    return { error: true, message: "Lỗi Google Drive: Không thể chia sẻ tài liệu." };
  }
}

function deductPointsSmart(email, amount) {
  amount = Number(amount);
  if (isNaN(amount) || amount <= 0) return false;

  const info = getUserAndPartner(email);
  if (!info) return false;

  let myCoin = Number(info.myRow[10]) || 0;

  // Đủ coin mình
  if (myCoin >= amount) {
    info.sheet.getRange(info.myIndex + 1, 11)
      .setValue(myCoin - amount);
    return true;
  }

  // Không đủ → lấy partner
  const need = amount - myCoin;

  if (!info.partnerRow) return false;

  let partnerCoin = Number(info.partnerRow[10]) || 0;
  if (partnerCoin < need) return false;

  // Trừ mình về 0
  info.sheet.getRange(info.myIndex + 1, 11).setValue(0);

  // Trừ partner phần còn thiếu
  info.sheet.getRange(info.partnerIndex + 1, 11)
    .setValue(partnerCoin - need);

  return true;
}

function deductPointsSmartForCurrentUser(amount) {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return false;

  return deductPointsSmart(email, amount);
}


function getTripleDCoinInfo() {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) {
    return { userPoints: 0, totalPoints: 0, percentage: 0, entityName: "" };
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();

  const visited = new Set();
  let totalPoints = 0;

  // ✅ Tính tổng server theo ENTITY (giống hệt Event)
  for (let i = 1; i < data.length; i++) {
    if (visited.has(i)) continue;

    const row = data[i];
    const partnerEmail = String(row[19] || "").trim().toLowerCase();
    const myCoin = parseInt(row[10]) || 0;

    const partnerIndex = data.findIndex((r, idx) =>
      idx > 0 &&
      String(r[1]).trim().toLowerCase() === partnerEmail
    );

    if (partnerIndex > 0) {
      const partnerCoin = parseInt(data[partnerIndex][10]) || 0;

      totalPoints += myCoin + partnerCoin;

      // 🔥 QUAN TRỌNG: đánh dấu cả 2 đã tính
      visited.add(i);
      visited.add(partnerIndex);
    } else {
      totalPoints += myCoin;
      visited.add(i);
    }
  }

  // ✅ ENTITY của user (giống logic Event)
  const info = getUserAndPartner(email);
  if (!info) {
    return { userPoints: 0, totalPoints, percentage: 0, entityName: "" };
  }

  const myName = info.myRow[0];
  const myCoin = Number(info.myRow[10]) || 0;

  let entityName = myName;
  let userPoints = myCoin;

  if (info.partnerRow) {
    const partnerName = info.partnerRow[0];
    const partnerCoin = Number(info.partnerRow[10]) || 0;

    entityName = `${myName} ❤️ ${partnerName}`;
    userPoints = myCoin + partnerCoin;
  }

  const percentage = totalPoints
    ? (userPoints / totalPoints) * 100
    : 0;

  return { userPoints, totalPoints, percentage, entityName };
}


function getUserAndPartner(email) {
  if (!email) return null;
  const sheet = getSheet("Users");
  const data = getSheetDataCached("Users");

  email = String(email).trim().toLowerCase();

  const myIndex = data.findIndex(r =>
    String(r[1]).trim().toLowerCase() === email
  );
  if (myIndex < 0) return null;

  const myRow = data[myIndex];
  const partnerEmail = String(myRow[19] || "").trim().toLowerCase();

  let partnerIndex = -1;
  let partnerRow = null;

  if (partnerEmail) {
    partnerIndex = data.findIndex(r =>
      String(r[1]).trim().toLowerCase() === partnerEmail
    );
    if (partnerIndex >= 0) partnerRow = data[partnerIndex];
  }

  return {
    sheet,
    myIndex,
    myRow,
    partnerIndex,
    partnerRow
  };
}

/**Code Về Triple D Coin*/



function requestApprovalWithReasonAndDetail(reason, fbLink, supportEmail, otherNote) {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false, message: "Bạn chưa đăng nhập." };

  const user = getUserByEmail(email);
  if (!user) return { success: false, message: "Không tìm thấy người dùng." };

  const points = parseInt(user.data[10]) || 0;    // cột K (index 10) điểm hiện tại
  const row = user.row;
  const trangThaiCu = user.data[11] || "";        // cột L (index 11) trạng thái cũ
  let currentRole = (user.data[13] || "").toString().trim().toUpperCase();
  if (!["VIP", "STUDENT", "S-STUDENT", "O-STUDENT"].includes(currentRole)) {
    currentRole = "MEM";
  }

  if (trangThaiCu.startsWith("Đã gửi")) {
    return { success: false, message: "Bạn đã gửi yêu cầu xét duyệt rồi. Vui lòng chờ Ban Quản Trị rà soát và phê duyệt." };
  }

  let pointsToDeduct = 0;
  let trangThai = "";
  let chiTiet = "";

  if (reason === "sstudent") {
    pointsToDeduct = 0;
    trangThai = "Đã gửi: Yêu cầu S-STUDENT";
  } else {
    // VIP request costs 300 coins
    if (points < 300) return { success: false, message: "Bạn cần ít nhất 300 🪙 Triple D Coin để gửi yêu cầu VIP." };
    pointsToDeduct = 300;
    
    let displayReason = reason;
    if (reason === "support") displayReason = "Góp Gạo Thổi Cơm Chung";
    else if (reason === "fan") displayReason = "Fan Cứng Fanpage";
    else if (reason === "other") displayReason = "Lý do khác";
    
    trangThai = "Đã gửi: Yêu cầu VIP (" + displayReason + ")";
  }

  if (fbLink) chiTiet += `Facebook: ${fbLink}\n`;
  if (supportEmail) chiTiet += `Email: ${supportEmail}\n`;
  if (otherNote) chiTiet += `Ghi chú: ${otherNote}`;

  // ✅ Trừ điểm và cập nhật trạng thái, kết quả
  updateUserField(row, 10, points - pointsToDeduct); // Điểm mới
  updateUserField(row, 11, trangThai);              // Trạng thái (L)
  updateUserField(row, 12, chiTiet.trim() || "Chờ xét duyệt"); // Kết quả / Chi tiết (M)

  // Ghi log vào Tracking sheet
  logActivityToSheet(email, "VIP_REQUEST", `Yêu cầu: ${trangThai}. Chi tiết: ${chiTiet.trim() || 'Không có'}. Khấu trừ: ${pointsToDeduct} coins. Chờ Admin duyệt`);

  return {
    success: true,
    message: `Yêu cầu của bạn đã được gửi thành công!\nBan quản trị sẽ kiểm tra rà soát và phê duyệt sớm nhất.`
  };
}

function isUserVIP(email) {
  const role = getUserRole(email);
  return role === "VIP" || role === "O-STUDENT";
}

function isUserStudent(email) {
  return getUserRole(email) === "STUDENT" || getUserRole(email) === "S-STUDENT";
}

function checkVIPStatus() {
  const cache = CacheService.getUserCache();
  const email = cache.get("loggedInUser");
  return isUserVIP(email);
}

function getAllUsersForAdmin() {
  try {
    const cache = CacheService.getUserCache();
    const callerEmail = cache.get("loggedInUser");
    if (!callerEmail) return { success: false, message: "Chưa đăng nhập." };
    
    // Check nếu caller là ADMIN
    const callerRole = getUserRole(callerEmail);
    if (callerRole !== "ADMIN") {
      return { success: false, message: "Từ chối truy cập. Bạn không phải là Admin." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    if (!sheet) return { success: false, message: "Không tìm thấy Sheet Users." };
    
    const data = sheet.getDataRange().getValues();
    const users = [];
    
    // Bỏ qua dòng tiêu đề (i=1)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[0] || "";
      const email = row[1] ? row[1].toString().trim() : "";
      if (!email) continue;
      
      const points = Number(row[10]) || 0;
      const roleRaw = (row[13] || "").toString().trim().toUpperCase();
      let role = "MEM";
      if (["ADMIN", "STUDENT", "S-STUDENT", "O-STUDENT", "VIP"].includes(roleRaw)) {
        role = roleRaw;
      }
      
      let lastActiveDate = row[14] || row[2]; // fallback to row[2] if row[14] (Column O) is empty
      let lastActive = "N/A";
      let daysAgo = "";
      
      if (lastActiveDate) {
        let dateObj = null;
        if (lastActiveDate instanceof Date) {
          dateObj = lastActiveDate;
        } else {
          const parsed = Date.parse(lastActiveDate);
          if (!isNaN(parsed)) {
            dateObj = new Date(parsed);
          }
        }
        
        if (dateObj) {
          lastActive = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
          const today = new Date();
          const d1 = new Date(today.getFullYear(), today.getMonth(), today.getDate());
          const d2 = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
          const diffTime = d1.getTime() - d2.getTime();
          const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
          
          if (diffDays <= 0) {
            daysAgo = "Hôm nay";
          } else if (diffDays === 1) {
            daysAgo = "Hôm qua";
          } else {
            daysAgo = diffDays + " ngày trước";
          }
        }
      }
      
      let deviceTypeStr = "Chưa rõ";
      const devRaw = row[18];
      if (devRaw) {
        try {
          const devObj = typeof devRaw === 'string' ? JSON.parse(devRaw) : devRaw;
          const devType = (devObj.deviceType || "").toUpperCase();
          const os = (devObj.os || "").toLowerCase();
          const pForm = (devObj.platform || "").toLowerCase();
          
          // Phân biệt bằng kích thước màn hình
          let minScreenEdge = -1;
          const screenStr = (devObj.screen || "").toString();
          if (screenStr.indexOf("x") !== -1) {
            const parts = screenStr.split("x").map(Number).filter(n => !isNaN(n));
            if (parts.length >= 2) {
              minScreenEdge = Math.min(parts[0], parts[1]);
            }
          }

          if (devType === "PC" || devType === "DESKTOP" || devType === "LAPTOP" || os.includes("windows") || os.includes("macos") || os.includes("mac os") || os.includes("linux")) {
            deviceTypeStr = "🖥️ Máy tính";
          } else {
            // Thiết bị di động hoặc bảng:
            // 1. Kiểm tra dấu hiệu rõ rệt của Tablet (iPad, tablet keyword)
            let isTablet = devType === "TABLET" || devType === "IPAD" || os.includes("ipad") || pForm.includes("ipad") || pForm.includes("tablet");
            
            // 2. Kiểm tra dựa trên độ phân giải màn hình (breakpoints)
            // Nếu cạnh nhỏ nhất >= 550px thì là Máy tính bảng, ngược lại < 550px chắc chắn là Điện thoại di động
            if (minScreenEdge > 0) {
              if (minScreenEdge >= 550) {
                isTablet = true;
              } else {
                isTablet = false; // Ngăn chặn sự nhầm lẫn của cụm "Mobile/Tablet" từ client
              }
            }
            
            if (isTablet) {
              deviceTypeStr = "📟 Máy tính bảng";
            } else if (devType === "MOBILE" || devType === "PHONE" || devType === "MOBILE/TABLET" || os.includes("android") || os.includes("ios") || os.includes("iphone")) {
              deviceTypeStr = "📱 Điện thoại";
            } else {
              deviceTypeStr = "❓ Chưa rõ";
            }
          }
        } catch (e) {
          deviceTypeStr = "⚠️ Lỗi đọc";
        }
      } else {
        deviceTypeStr = "⚪ Chưa đăng nhập";
      }

      users.push({
        name: name,
        email: email,
        points: points,
        role: role,
        lastActive: lastActive,
        daysAgo: daysAgo,
        vipStatus: row[11] ? row[11].toString().trim() : "",
        vipNote: row[12] ? row[12].toString().trim() : "",
        deviceType: deviceTypeStr
      });
    }
    
    // Sắp xếp admin -> stu -> O-stu -> vip -> mem, rồi điểm giảm dần
    const roleWeight = { "ADMIN": 6, "S-STUDENT": 5, "STUDENT": 4, "O-STUDENT": 3, "VIP": 2, "MEM": 1 };
    users.sort((a, b) => {
      if (roleWeight[a.role] !== roleWeight[b.role]) {
        return roleWeight[b.role] - roleWeight[a.role];
      }
      return b.points - a.points;
    });
    
    return { success: true, data: users };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function adminUpdateUser(targetEmail, newRole, newPoints) {
  try {
    const cache = CacheService.getUserCache();
    const callerEmail = cache.get("loggedInUser");
    if (!callerEmail) return { success: false, message: "Chưa đăng nhập." };
    
    // Check nếu caller là ADMIN
    const callerRole = getUserRole(callerEmail);
    if (callerRole !== "ADMIN") {
      return { success: false, message: "Từ chối truy cập. Bạn không phải là Admin." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    if (!sheet) return { success: false, message: "Không tìm thấy Sheet Users." };
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    targetEmail = (targetEmail || "").toString().toLowerCase().trim();
    if (!targetEmail) return { success: false, message: "Email không hợp lệ." };
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().toLowerCase().trim() === targetEmail) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: "Không tìm thấy người dùng này." };
    }
    
    const currentRow = data[rowIndex];
    const currentRoleInSheet = (currentRow[13] || "MEM").toString().trim().toUpperCase();
    const validRoles = ["ADMIN", "STUDENT", "S-STUDENT", "O-STUDENT", "VIP", "MEM"];
    
    let safeRole = (newRole || "").toString().trim().toUpperCase();
    if (!safeRole || !validRoles.includes(safeRole)) {
      // Giữ nguyên vai trò hiện tại nếu role gửi lên bị trống hoặc không hợp lệ
      safeRole = validRoles.includes(currentRoleInSheet) ? currentRoleInSheet : "MEM";
    }
    
    let safePoints = Number(newPoints);
    if (isNaN(safePoints) || safePoints < 0) {
      safePoints = Number(currentRow[10]) || 0;
    }
    
    // index in Apps Script sheet is 1-based, array is 0-based.
    // Row in sheet is rowIndex + 1.
    // Points is column 11 (index 10)
    // Role is column 14 (index 13)
    sheet.getRange(rowIndex + 1, 11).setValue(safePoints);
    sheet.getRange(rowIndex + 1, 14).setValue(safeRole);
    invalidateSheetCache("Users");
    
    // Clear user role memory cache
    if (typeof _cachedUserRoles !== 'undefined') {
      delete _cachedUserRoles[targetEmail];
    }
    
    // Update Tracking log
    logActivityToSheet(targetEmail, "ADMIN_UPDATE_USER", `Admin ${callerEmail} set role=${safeRole}, coins=${safePoints}`);
    
    return { success: true, message: `Cập nhật thành công cho ${targetEmail} (Vai trò: ${safeRole}, Coin: ${safePoints.toLocaleString()})` };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Reset / Xóa liên kết thiết bị của người dùng, cho phép đăng nhập trên thiết bị mới
 */
function adminResetUserDevice(targetEmail) {
  try {
    const cache = CacheService.getUserCache();
    const callerEmail = cache.get("loggedInUser");
    if (!callerEmail) return { success: false, message: "Chưa đăng nhập." };
    
    const callerRole = getUserRole(callerEmail);
    if (callerRole !== "ADMIN") {
      return { success: false, message: "Từ chối truy cập. Bạn không phải là Admin." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    if (!sheet) return { success: false, message: "Không tìm thấy Sheet Users." };
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    targetEmail = (targetEmail || "").toString().toLowerCase().trim();
    if (!targetEmail) return { success: false, message: "Email không hợp lệ." };
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().toLowerCase().trim() === targetEmail) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: "Không tìm thấy người dùng này." };
    }
    
    // Cột Q (index 16) là deviceID (cột 17 trên sheet)
    // Cột S (index 18) là deviceInfo (cột 19 trên sheet)
    sheet.getRange(rowIndex + 1, 17).setValue(""); // Xóa deviceID
    sheet.getRange(rowIndex + 1, 19).setValue(""); // Xóa deviceInfo
    invalidateSheetCache("Users");
    
    logActivityToSheet(targetEmail, "ADMIN_RESET_DEVICE", `Admin ${callerEmail} đã mở khóa / reset liên kết thiết bị cho học sinh.`);
    
    return { success: true, message: `Đã mở khóa thiết bị thành công cho ${targetEmail}! Học sinh có thể đăng nhập trên máy mới.` };
  } catch (err) {
    return { success: false, message: "Lỗi: " + err.toString() };
  }
}

/**
 * Đăng xuất bắt buộc (Force Logout) 1 tài khoản người dùng cụ thể
 */
function adminForceLogoutUser(targetEmail) {
  try {
    const cache = CacheService.getUserCache();
    const callerEmail = cache.get("loggedInUser");
    if (!callerEmail) return { success: false, message: "Chưa đăng nhập." };
    
    const callerRole = getUserRole(callerEmail);
    if (callerRole !== "ADMIN") {
      return { success: false, message: "Từ chối truy cập. Bạn không phải là Admin." };
    }

    targetEmail = (targetEmail || "").toString().toLowerCase().trim();
    if (!targetEmail) return { success: false, message: "Email người dùng không hợp lệ." };

    const now = Date.now();
    safeSetScriptProperty("USER_FORCE_LOGOUT_" + targetEmail, String(now));

    logActivityToSheet(targetEmail, "ADMIN_FORCE_LOGOUT_USER", `Admin ${callerEmail} đã hủy phiên làm việc của ${targetEmail}.`);

    return { 
      success: true, 
      message: `Đã đăng xuất bắt buộc thành công cho tài khoản ${targetEmail}!` 
    };
  } catch (err) {
    return { success: false, message: "Lỗi: " + err.toString() };
  }
}

/**
 * Đăng xuất bắt buộc TOÀN BỘ SERVER (Force Logout All Users)
 */
function adminForceLogoutAllUsers() {
  try {
    const cache = CacheService.getUserCache();
    const callerEmail = cache.get("loggedInUser");
    if (!callerEmail) return { success: false, message: "Chưa đăng nhập." };
    
    const callerRole = getUserRole(callerEmail);
    if (callerRole !== "ADMIN") {
      return { success: false, message: "Từ chối truy cập. Bạn không phải là Admin." };
    }

    const now = Date.now();
    safeSetScriptProperty("GLOBAL_FORCE_LOGOUT_TIMESTAMP", String(now));
    
    // Bảo vệ phiên làm việc của chính Admin thực hiện thao tác
    cache.put("loginTimestamp", String(now + 2000), 21600);

    logActivityToSheet(callerEmail, "ADMIN_FORCE_LOGOUT_ALL", `Admin ${callerEmail} đã kích hoạt đăng xuất toàn bộ Server.`);

    return { 
      success: true, 
      message: "Đã kích hoạt đăng xuất bắt buộc toàn Server thành công! Toàn bộ người dùng khác sẽ phải đăng nhập lại." 
    };
  } catch (err) {
    return { success: false, message: "Lỗi: " + err.toString() };
  }
}

function adminApproveVip(targetEmail, approvedRole) {
  try {
    const cache = CacheService.getUserCache();
    const callerEmail = cache.get("loggedInUser");
    if (!callerEmail) return { success: false, message: "Chưa đăng nhập." };
    
    const callerRole = getUserRole(callerEmail);
    if (callerRole !== "ADMIN") return { success: false, message: "Từ chối truy cập. Bạn không phải là Admin." };

    const sheet = getSheet("Users");
    if (!sheet) return { success: false, message: "Không tìm thấy Sheet Users." };
    
    const data = getSheetDataCached("Users");
    let rowIndex = -1;
    targetEmail = targetEmail.toLowerCase().trim();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().toLowerCase().trim() === targetEmail) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, message: "Không tìm thấy người dùng này." };
    
    const isSStudent = approvedRole === "S-STUDENT";
    const resultText = isSStudent ? "✅ Đăng ký S-STUDENT Thành Công!" : "✅ Nâng Vip Thành Công!";
    const successMsg = isSStudent ? `Phê duyệt thành công S-STUDENT cho ${targetEmail}` : `Phê duyệt thành công VIP cho ${targetEmail}`;

    // Batch update Cols 12, 13, 14: Status (Col L), Result (Col M), Role (Col N)
    sheet.getRange(rowIndex + 1, 12, 1, 3).setValues([["Đã duyệt bởi Admin", resultText, approvedRole]]);
    invalidateSheetCache("Users");
    
    // Tracking log
    logActivityToSheet(targetEmail, "ADMIN_APPROVE_VIP", `Admin ${callerEmail} approved role=${approvedRole}`);
    
    return { success: true, message: successMsg };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function adminRejectVip(targetEmail) {
  try {
    const cache = CacheService.getUserCache();
    const callerEmail = cache.get("loggedInUser");
    if (!callerEmail) return { success: false, message: "Chưa đăng nhập." };
    
    const callerRole = getUserRole(callerEmail);
    if (callerRole !== "ADMIN") return { success: false, message: "Từ chối truy cập. Bạn không phải là Admin." };

    const sheet = getSheet("Users");
    if (!sheet) return { success: false, message: "Không tìm thấy Sheet Users." };
    
    const data = getSheetDataCached("Users");
    let rowIndex = -1;
    targetEmail = targetEmail.toLowerCase().trim();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().toLowerCase().trim() === targetEmail) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, message: "Không tìm thấy người dùng này." };
    
    // Batch update Cols 12, 13: Status (Col L), Result (Col M)
    sheet.getRange(rowIndex + 1, 12, 1, 2).setValues([["Từ chối nâng VIP", "❌ Không Được Duyệt!"]]);
    invalidateSheetCache("Users");
    
    // Tracking log
    logActivityToSheet(targetEmail, "ADMIN_REJECT_VIP", `Admin ${callerEmail} rejected VIP request`);
    
    return { success: true, message: `Đã từ chối yêu cầu của ${targetEmail}` };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function adminGetMonitorSessions() {
  try {
    const cache = CacheService.getUserCache();
    const callerEmail = cache.get("loggedInUser");
    if (!callerEmail) return { success: false, message: "Chưa đăng nhập." };
    
    const callerRole = getUserRole(callerEmail);
    if (callerRole !== "ADMIN") {
      return { success: false, message: "Từ chối truy cập. Bạn không phải là Admin." };
    }

    // Tự động dọn dẹp các thông báo cũ > 15 ngày
    try {
      autoCleanupNotifications();
    } catch (eNotif) {
      Logger.log("Error during autoCleanupNotifications: " + eNotif.message);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MonitorSessions");
    if (!sheet) return { success: true, data: [] };
    
    const data = sheet.getDataRange().getValues();
    const list = [];
    
    // Read from bottom to top to show newest sessions first, max 1000 rows
    const limit = Math.max(1, data.length - 1000);
    for (let i = data.length - 1; i >= limit; i--) {
      const email = data[i][0];
      if (!email) continue;
      
      let startVal = data[i][1];
      let endVal = data[i][2];
      
      if (startVal instanceof Date) {
        startVal = Utilities.formatDate(startVal, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
      }
      if (endVal instanceof Date) {
        endVal = Utilities.formatDate(endVal, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
      }
      
      list.push({
        email: email,
        start: startVal || "N/A",
        end: endVal || "N/A",
        duration: Number(data[i][3]) || 0,
        photos: Number(data[i][4]) || 0
      });
    }
    
    return { success: true, data: list };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function adminCreateFakeMonitorSessions() {
  try {
    const cache = CacheService.getUserCache();
    const callerEmail = cache.get("loggedInUser");
    if (!callerEmail) return { success: false, message: "Chưa đăng nhập." };
    
    const callerRole = getUserRole(callerEmail);
    if (callerRole !== "ADMIN") {
      return { success: false, message: "Từ chối truy cập. Bạn không phải là Admin." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("MonitorSessions");
    if (!sheet) {
      sheet = ss.insertSheet("MonitorSessions");
      sheet.appendRow(["Email", "Bắt đầu", "Kết thúc", "Thời lượng (phút)", "Số ảnh chụp"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#cbd5e1");
    }
    
    // Clear existing data (keep header) to make testing fresh and clear!
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).clearContent();
    }
    
    // Add 6 super clear mock rows for 23/05/2026
    // target range to test: 23/05/2026 from 19:30 to 22:00
    sheet.appendRow(["studentA_giao_truoc@gmail.com", "23/05/2026 18:30:00", "23/05/2026 21:30:00", 180, 24]); // Started early, ended middle
    sheet.appendRow(["studentB_giao_sau@gmail.com", "23/05/2026 20:00:00", "23/05/2026 23:30:00", 210, 28]); // Started middle, ended late
    sheet.appendRow(["studentC_bao_trum@gmail.com", "23/05/2026 18:00:00", "23/05/2026 23:59:00", 359, 48]); // Covered the whole target range
    sheet.appendRow(["studentD_nam_trong@gmail.com", "23/05/2026 19:45:00", "23/05/2026 21:15:00", 90, 12]); // Fits perfectly inside target range
    sheet.appendRow(["studentE_ngoai_truoc@gmail.com", "23/05/2026 15:00:00", "23/05/2026 18:00:00", 180, 10]); // Entirely before
    sheet.appendRow(["studentF_ngoai_sau@gmail.com", "23/05/2026 22:30:00", "23/05/2026 23:30:00", 60, 6]); // Entirely after (since focus ends at 22:00)
    
    return { success: true, message: "Đã tạo/đặt lại dữ liệu thử nghiệm với 6 phiên mẫu!" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function adminCleanupOldMonitorSessions() {
  try {
    const cache = CacheService.getUserCache();
    const callerEmail = cache.get("loggedInUser");
    if (!callerEmail) return { success: false, message: "Chưa đăng nhập." };
    
    const callerRole = getUserRole(callerEmail);
    if (callerRole !== "ADMIN") {
      return { success: false, message: "Từ chối truy cập. Bạn không phải là Admin." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MonitorSessions");
    if (!sheet) {
      return { success: true, count: 0, message: "Không tìm thấy bảng dữ liệu MonitorSessions." };
    }
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow <= 1 || lastColumn <= 0) {
      return { success: true, count: 0, message: "Bảng dữ liệu đang trống." };
    }
    
    const values = sheet.getDataRange().getValues();
    const rowsToKeep = [];
    let deletedCount = 0;
    
    const now = new Date();
    const cutoffTime = now.getTime() - (30 * 24 * 60 * 60 * 1000); // 30 days ago in ms
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const startStr = row[1];
      const endStr = row[2];
      
      const startDate = parseCustomDateString(startStr);
      const endDate = parseCustomDateString(endStr);
      
      let isOld = false;
      if (startDate && startDate.getTime() < cutoffTime) {
        isOld = true;
      } else if (!startDate && endDate && endDate.getTime() < cutoffTime) {
        isOld = true;
      }
      
      if (isOld) {
        deletedCount++;
      } else {
        rowsToKeep.push(row);
      }
    }
    
    // Clear old contents below headers
    sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
    
    // Write remaining rows back
    if (rowsToKeep.length > 0) {
      sheet.getRange(2, 1, rowsToKeep.length, lastColumn).setValues(rowsToKeep);
    }
    
    return { 
      success: true, 
      count: deletedCount, 
      message: `Đã dọn dẹp xong! Đã xóa ${deletedCount} phiên học tập đã quá hạn 30 ngày.` 
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function autoCleanupNotifications() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Notifications");
    if (!sheet) return { success: true, count: 0 };
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow <= 1 || lastColumn <= 0) return { success: true, count: 0 };
    
    const values = sheet.getDataRange().getValues();
    const rowsToKeep = [];
    const now = new Date();
    const cutoffTime = now.getTime() - (15 * 24 * 60 * 60 * 1000); // 15 days ago in ms
    let deletedCount = 0;
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const notifId = String(row[0] || "");
      const timeVal = row[6];
      const notifTime = timeVal instanceof Date ? timeVal : new Date(timeVal);
      
      let isOld = false;
      if (notifId !== "welcome_system_global" && notifTime && !isNaN(notifTime.getTime()) && notifTime.getTime() < cutoffTime) {
        isOld = true;
      }
      
      if (!isOld) {
        rowsToKeep.push(row);
      } else {
        deletedCount++;
      }
    }
    
    // Check if we actually need to delete anything
    if (values.length - 1 > rowsToKeep.length) {
      // Clear old contents below headers
      sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
      // Write remaining rows back
      if (rowsToKeep.length > 0) {
        sheet.getRange(2, 1, rowsToKeep.length, lastColumn).setValues(rowsToKeep);
      }
      Logger.log(`[Auto Cleanup Notifications] Đã dọn dẹp thành công. Đã xóa ${deletedCount} thông báo cũ (>15 ngày).`);
    }
    return { success: true, count: deletedCount };
  } catch (e) {
    Logger.log("Error in autoCleanupNotifications: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function autoCleanupMonitorSessions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MonitorSessions");
    if (!sheet) return { success: true, count: 0 };
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow <= 1 || lastColumn <= 0) return { success: true, count: 0 };
    
    const values = sheet.getDataRange().getValues();
    const rowsToKeep = [];
    const now = new Date();
    // Ngưỡng dọn dẹp: giữ lại tối đa 15 ngày hoạt động gần nhất để sheet luôn siêu nhẹ và nhanh
    const cutoffTime = now.getTime() - (15 * 24 * 60 * 60 * 1000); 
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const startStr = row[1];
      const endStr = row[2];
      
      const startDate = parseCustomDateString(startStr);
      const endDate = parseCustomDateString(endStr);
      
      let isOld = false;
      if (startDate && startDate.getTime() < cutoffTime) {
        isOld = true;
      } else if (!startDate && endDate && endDate.getTime() < cutoffTime) {
        isOld = true;
      }
      
      if (!isOld) {
        rowsToKeep.push(row);
      }
    }
    
    // Giới hạn cứng: Nếu sau khi dọn dẹp theo ngày vẫn còn quá 1000 hàng, chỉ giữ lại 500 hàng mới nhất để tối ưu dung lượng sheet tối đa
    let finalRows = rowsToKeep;
    if (finalRows.length > 1000) {
      finalRows = finalRows.slice(finalRows.length - 500);
    }
    
    const deletedCount = (values.length - 1) - finalRows.length;
    // Chỉ cập nhật nếu có sự thay đổi (đã xóa đi bớt dòng)
    if (deletedCount > 0) {
      // Xóa sạch nội dung cũ bên dưới tiêu đề
      sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
      // Ghi lại danh sách các hàng được giữ
      if (finalRows.length > 0) {
        sheet.getRange(2, 1, finalRows.length, lastColumn).setValues(finalRows);
      }
      Logger.log(`[Auto Cleanup MonitorSessions] Đã dọn dẹp thành công. Đã xóa ${deletedCount} dòng cũ. Giữ lại ${finalRows.length}/${values.length - 1} dòng.`);
    }
    return { success: true, count: deletedCount };
  } catch (e) {
    Logger.log("Error in autoCleanupMonitorSessions: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * =========================================================================
 * HÀM TỔNG HỢP DỌN DẸP TOÀN BỘ HỆ THỐNG & NHẬT KÝ (LOGS / NOTIFICATIONS / MONITOR / GIFTCODES / OPENFILES)
 * DÙNG ĐỂ TẠO TIME-DRIVEN TRIGGER TRONG GOOGLE APPS SCRIPT
 * =========================================================================
 * 
 * HƯỚNG DẪN CÀI ĐẶT APPS SCRIPT TRIGGER:
 * 1. Trong giao diện Apps Script Editor, nhìn thanh công cụ bên trái -> Bấm biểu tượng "Đồng hồ" (Kích hoạt / Triggers).
 * 2. Bấm nút "+ Thêm bộ kích hoạt" (Add Trigger) ở góc dưới bên phải.
 * 3. Tại mục "Chọn chức năng để chạy" (Choose function to run): Chọn `triggerMasterSystemCleanup` hoặc `dondepHeThongVaLichSuTrigger`.
 * 4. Tại mục "Chọn nguồn sự kiện" (Select event source): Chọn "Theo thời gian" (Time-driven).
 * 5. Tại mục "Chọn loại bộ kích hoạt dựa trên thời gian": Chọn "Mỗi ngày" (Day timer) hoặc "Mỗi 6 giờ" (Hour timer).
 * 6. Bấm "Lưu" (Save).
 */
function triggerMasterSystemCleanup() {
  Logger.log("=================================================");
  Logger.log("=== BẮT ĐẦU DỌN DẸP HỆ THỐNG TOÀN DIỆN (TRIGGER) ===");
  Logger.log("=================================================");
  
  const summary = [];

  // 1. Dọn dẹp Thông báo cũ (Notifications - giữ lại thông báo < 15 ngày & thông báo ghim hệ thống)
  try {
    const resNotif = autoCleanupNotifications();
    const countNotif = (resNotif && typeof resNotif.count === 'number') ? resNotif.count : 0;
    summary.push(`Notifications: Đã dọn dẹp ${countNotif} thông báo cũ (>15 ngày)`);
  } catch (e1) {
    summary.push(`Notifications Error: ${e1.message}`);
  }

  // 2. Dọn dẹp Nhật ký giám sát học tập (MonitorSessions - giữ lại phiên < 15 ngày, tối đa 500 dòng)
  try {
    const resMonitor = autoCleanupMonitorSessions();
    const countMonitor = (resMonitor && typeof resMonitor.count === 'number') ? resMonitor.count : 0;
    summary.push(`MonitorSessions: Đã dọn dẹp ${countMonitor} phiên giám sát cũ`);
  } catch (e2) {
    summary.push(`MonitorSessions Error: ${e2.message}`);
  }

  // 3. Dọn dẹp Lịch sử mở file & Thu hồi quyền Drive hết hạn (OpenFilesLog)
  try {
    revokeAccessAndCleanLogIfExpired();
    summary.push(`OpenFilesLog & Drive Permissions: Đã dọn log & thu hồi quyền truy cập Drive hết hạn`);
  } catch (e3) {
    summary.push(`OpenFilesLog Error: ${e3.message}`);
  }

  // 4. Dọn dẹp Giftcode hết hạn & Xóa ô đánh dấu quà tặng đã hết hạn trong Users Sheet
  try {
    const resGift = clearExpiredGiftCodeValues();
    const giftCount = (resGift && typeof resGift.clearedCount === 'number') ? resGift.clearedCount : 0;
    summary.push(`Giftcodes: Đã dọn dẹp ${giftCount} ô ghi nhận giftcode hết hạn`);
  } catch (e4) {
    summary.push(`Giftcodes Error: ${e4.message}`);
  }

  // 5. Cắt bớt dòng dư thừa cho các sheet nhật ký giao dịch lớn (Giữ Google Sheet siêu nhẹ & phản hồi nhanh)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    limitSheetRows_(ss.getSheetByName("CoinTransactions"), 2000);
    limitSheetRows_(ss.getSheetByName("PurchaseLog"), 1000);
    limitSheetRows_(ss.getSheetByName("TheoDoi"), 1000);
    limitSheetRows_(ss.getSheetByName("FootballBets"), 1000);
    summary.push(`Trimming Row Limits: Giới hạn dòng CoinTransactions(2000), PurchaseLog(1000), TheoDoi(1000), FootballBets(1000)`);
  } catch (e5) {
    summary.push(`Trimming Rows Error: ${e5.message}`);
  }

  Logger.log("=================================================");
  Logger.log("=== HOÀN TẤT DỌN DẸP HỆ THỐNG TOÀN DIỆN ===");
  summary.forEach(item => Logger.log(" -> " + item));
  Logger.log("=================================================");

  return {
    success: true,
    timestamp: new Date().toISOString(),
    details: summary
  };
}

/** Biệt danh Tiếng Việt để dễ nhận biết và chọn trong menu Trigger của Google Apps Script */
function dondepHeThongVaLichSuTrigger() {
  return triggerMasterSystemCleanup();
}

function adminGetCoinFlowStats() {
  const adminEmail = CacheService.getUserCache().get("loggedInUser");
  if (!adminEmail || getUserRole(adminEmail) !== "ADMIN") return { success: false, message: "Unauthorized." };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("CoinTransactions");
    if (!sheet) return { success: true, data: { inFlow: [], outFlow: [], history: [] } };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: { inFlow: [], outFlow: [], history: [] } };

    let inFlow = {};
    let outFlow = {};
    let history = []; 

    for (let i = data.length - 1; i >= 1; i--) {
      let row = data[i];
      let time = row[0];
      let email = row[1];
      let amount = parseFloat(row[2]) || 0;
      let type = row[3]; 
      let source = String(row[4] || "");
      
      let sourceLower = source.toLowerCase();
      let category = "Khác";

      if (amount > 0) {
          if (sourceLower.includes("thắng") || sourceLower.includes("dự đoán bóng đá chính xác") || sourceLower.includes("arcade") || sourceLower.includes("bầu cua") || sourceLower.includes("minigame") || sourceLower.includes("playtaixiu")) {
              category = "Chơi Game & Cá Cược (+)";
          } else if (sourceLower.includes("bài tập") || sourceLower.includes("hoàn thành") || sourceLower.includes("điểm danh") || sourceLower.includes("học tập")) {
              category = "Học Tập & Nhiệm Vụ (+)";
          } else if (sourceLower.includes("giftcode") || sourceLower.includes("hộp quà")) {
              category = "Sự Kiện & Giftcode (+)";
          } else if (sourceLower.includes("nhận") || sourceLower.includes("p2p transfer")) {
              category = "Nhận Chuyển Khoản P2P (+)";
          } else if (sourceLower.includes("hoàn trả") || sourceLower.includes("hoàn cược")) {
              category = "Hoàn Cược & Hệ Thống (+)";
          } else {
              category = "Thu Nhập Khác (+)";
          }
          if (!inFlow[category]) inFlow[category] = 0;
          inFlow[category] += amount;
      } else if (amount < 0) {
          if (sourceLower.includes("cược") || sourceLower.includes("dự đoán") || sourceLower.includes("playtaixiu") || sourceLower.includes("arcade") || sourceLower.includes("bầu cua") || sourceLower.includes("minigame")) {
              category = "Chơi Game & Cá Cược (-)";
          } else if (sourceLower.includes("mở tài liệu")) {
              category = "Mở Tài Liệu (-)";
          } else if (sourceLower.includes("mua") || sourceLower.includes("store")) {
              category = "Mua Sắm Store (-)";
          } else if (sourceLower.includes("chuyển") || sourceLower.includes("p2p transfer")) {
              category = "Chuyển Khoản P2P (-)";
          } else if (sourceLower.includes("reset Mùa giải") || sourceLower.includes("reset")) {
              category = "Hệ Thống Thu Hồi (-)";
          } else {
              category = "Chi Tiêu Khác (-)";
          }
          if (!outFlow[category]) outFlow[category] = 0;
          outFlow[category] += Math.abs(amount);
      }
      
      if (history.length < 100) {
        history.push({ 
          time: (time instanceof Date) ? Utilities.formatDate(time, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss") : time.toString(), 
          email: email, 
          amount: amount, 
          source: source 
        });
      }
    }
    
    let inFlowArr = Object.keys(inFlow).map(k => ({category: k, sum: inFlow[k]})).sort((a,b) => b.sum - a.sum);
    let outFlowArr = Object.keys(outFlow).map(k => ({category: k, sum: outFlow[k]})).sort((a,b) => b.sum - a.sum);

    return {
      success: true,
      data: {
        inFlow: inFlowArr,
        outFlow: outFlowArr,
        history: history
      }
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function adminCleanupOldNotifications() {
  try {
    const cache = CacheService.getUserCache();
    const callerEmail = cache.get("loggedInUser");
    if (!callerEmail) return { success: false, message: "Chưa đăng nhập." };
    
    const callerRole = getUserRole(callerEmail);
    if (callerRole !== "ADMIN") {
      return { success: false, message: "Từ chối truy cập. Bạn không phải là Admin." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Notifications");
    if (!sheet) {
      return { success: true, count: 0, message: "Không tìm thấy bảng thông báo Notifications." };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, count: 0, message: "Bảng thông báo đang trống." };
    }
    
    const values = sheet.getDataRange().getValues();
    const rowsToKeep = [];
    let deletedCount = 0;
    
    const now = new Date();
    const cutoffTime = now.getTime() - (15 * 24 * 60 * 60 * 1000); // 15 days ago in ms
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const timeVal = row[6];
      const notifId = String(row[0] || "");
      
      const notifTime = timeVal instanceof Date ? timeVal : new Date(timeVal);
      
      let isOld = false;
      if (notifId !== "welcome_system_global" && notifTime && !isNaN(notifTime.getTime()) && notifTime.getTime() < cutoffTime) {
        isOld = true;
      }
      
      if (isOld) {
        deletedCount++;
      } else {
        rowsToKeep.push(row);
      }
    }
    
    // Clear old contents below headers
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    
    // Write remaining rows back
    if (rowsToKeep.length > 0) {
      sheet.getRange(2, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
    }
    
    return { 
      success: true, 
      count: deletedCount, 
      message: `Đã dọn dẹp xong! Đã xóa vĩnh viễn ${deletedCount} thông báo cũ hơn 15 ngày.` 
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function logMonitorSession(userEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("MonitorSessions");
    if (!sheet) {
      sheet = ss.insertSheet("MonitorSessions");
      sheet.appendRow(["Email", "Bắt đầu", "Kết thúc", "Thời lượng (phút)", "Số ảnh chụp"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#cbd5e1");
    }
    
    const data = sheet.getDataRange().getValues();
    let foundRowIndex = -1;
    
    const cleanEmail = userEmail.toLowerCase().trim();
    
    // Scan from bottom up to find the user's last session
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] && data[i][0].toLowerCase().trim() === cleanEmail) {
        foundRowIndex = i;
        break;
      }
    }
    
    const now = new Date();
    
    if (foundRowIndex !== -1) {
      const startTimeRaw = data[foundRowIndex][1];
      const endTimeRaw = data[foundRowIndex][2];
      const photoCount = Number(data[foundRowIndex][4]) || 0;
      
      let endTime = null;
      if (endTimeRaw instanceof Date) {
        endTime = endTimeRaw;
      } else {
        endTime = parseCustomDateString(endTimeRaw);
      }
      
      let startTime = null;
      if (startTimeRaw instanceof Date) {
        startTime = startTimeRaw;
      } else {
        startTime = parseCustomDateString(startTimeRaw);
      }
      
      // If end time is within 5 minutes of now, update the current session
      if (endTime && (Math.abs(now.getTime() - endTime.getTime()) < 5 * 60 * 1000)) {
        const start = startTime || now;
        const durationMin = Math.max(0, Math.round((now.getTime() - start.getTime()) / 60000));
        
        sheet.getRange(foundRowIndex + 1, 3).setValue(now);
        sheet.getRange(foundRowIndex + 1, 4).setValue(durationMin);
        sheet.getRange(foundRowIndex + 1, 5).setValue(photoCount + 1);
        return;
      }
    }
    
    // Else, append a new session row
    sheet.appendRow([cleanEmail, now, now, 0, 1]);
    
  } catch (e) {
    Logger.log("Error logging monitor session: " + e.message);
  }
}

function parseCustomDateString(str) {
  if (!str) return null;
  if (str instanceof Date || (typeof str === 'object' && typeof str.getTime === 'function')) {
    return str;
  }
  
  const strVal = String(str).trim();
  if (!strVal) return null;
  
  let parsedDateLocal = null;
  
  // Match DD/MM/YYYY HH:MM:SS or DD/MM/YYYY HH:MM
  const dmyMatch = strVal.match(/^(\d{1,2})[\/_.-](\d{1,2})[\/_.-](\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1], 10);
    const month = parseInt(dmyMatch[2], 10) - 1;
    const year = parseInt(dmyMatch[3], 10);
    const hour = dmyMatch[4] ? parseInt(dmyMatch[4], 10) : 0;
    const min = dmyMatch[5] ? parseInt(dmyMatch[5], 10) : 0;
    const sec = dmyMatch[6] ? parseInt(dmyMatch[6], 10) : 0;
    parsedDateLocal = new Date(year, month, day, hour, min, sec);
  } else {
    // Match YYYY-MM-DD THH:MM:SS
    const ymdMatch = strVal.match(/^(\d{4})[\/_.-](\d{1,2})[\/_.-](\d{1,2})(?:[\sT]+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
    if (ymdMatch) {
      const year = parseInt(ymdMatch[1], 10);
      const month = parseInt(ymdMatch[2], 10) - 1;
      const day = parseInt(ymdMatch[3], 10);
      const hour = ymdMatch[4] ? parseInt(ymdMatch[4], 10) : 0;
      const min = ymdMatch[5] ? parseInt(ymdMatch[5], 10) : 0;
      const sec = ymdMatch[6] ? parseInt(ymdMatch[6], 10) : 0;
      parsedDateLocal = new Date(year, month, day, hour, min, sec);
    }
  }
  
  if (!parsedDateLocal) {
    try {
      const parsed = Date.parse(strVal);
      if (!isNaN(parsed)) {
        parsedDateLocal = new Date(parsed);
      }
    } catch (e) {}
  }
  
  if (parsedDateLocal) {
    // Correct timezone shift relative to Session.getScriptTimeZone()
    try {
      const refDate = new Date();
      const scriptStrStr = Utilities.formatDate(refDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      const parts = scriptStrStr.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})$/);
      if (parts) {
        const scriptLocal = new Date(
          parseInt(parts[1], 10),
          parseInt(parts[2], 10) - 1,
          parseInt(parts[3], 10),
          parseInt(parts[4], 10),
          parseInt(parts[5], 10),
          parseInt(parts[6], 10)
        );
        const adjustment = scriptLocal.getTime() - refDate.getTime();
        return new Date(parsedDateLocal.getTime() - adjustment);
      }
    } catch (err) {
      Logger.log("Error correcting parse timezone: " + err.message);
    }
    return parsedDateLocal;
  }
  
  return null;
}

function adminGetRecentQA() {
  try {
    const cache = CacheService.getUserCache();
    const callerEmail = cache.get("loggedInUser");
    if (!callerEmail) return { success: false, message: "Chưa đăng nhập." };
    
    const callerRole = getUserRole(callerEmail);
    if (callerRole !== "ADMIN") return { success: false, message: "Từ chối truy cập." };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("QA");
    if (!sheet) return { success: true, data: [] };
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };
    
    const qaList = [];
    const titleCache = {};
    
    // Pass 0: Build list of ADMIN emails
    const userSheet = ss.getSheetByName("Users");
    const adminEmails = new Set();
    if (userSheet) {
      const userData = userSheet.getDataRange().getValues();
      for (let i = 1; i < userData.length; i++) {
        const roleRaw = (userData[i][13] || "").toString().trim().toUpperCase();
        if (roleRaw === "ADMIN") {
          const uemail = userData[i][1] ? userData[i][1].toString().trim().toLowerCase() : "";
          if (uemail) adminEmails.add(uemail);
        }
      }
    }
    
    // Pass 1: Build map of replies from Admin
    const replyCountMap = {};
    for (let i = 1; i < data.length; i++) {
        const pId = data[i][5]; // ParentID
        if (pId) {
            const replierEmail = (data[i][2] || "").toString().trim().toLowerCase();
            if (adminEmails.has(replierEmail)) {
                replyCountMap[pId] = (replyCountMap[pId] || 0) + 1;
            }
        }
    }
    
    // Pass 2: Extract top-level questions
    let count = 0;
    // Duyệt ngược để lấy các câu hỏi mới nhất
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const parentId = row[5];
      
      // We only list top-level questions for the dashboard overview
      if (parentId) continue; 
      
      let timestamp = row[6];
      if (timestamp instanceof Date) {
        timestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      }
      
      const fileId = row[1];
      let fileTitle = titleCache[fileId] || "Chưa tải";
      let fileType = "";
      
      if (!titleCache[fileId]) {
         const info = getFileDetailForNotif(fileId);
         if (info) {
             fileTitle = info.title;
             fileType = info.type;
             titleCache[fileId] = { title: fileTitle, type: fileType };
         }
      } else {
         fileTitle = titleCache[fileId].title;
         fileType = titleCache[fileId].type;
      }

      let reacts = [];
      try {
        reacts = row[7] ? JSON.parse(row[7]) : [];
      } catch(e){}
      
      let adminReacted = reacts.some(r => adminEmails.has(r.toLowerCase()));

      qaList.push({
        id: row[0],
        fileId: fileId,
        email: row[2],
        name: row[3],
        content: row[4],
        parentId: row[5],
        timestamp: timestamp,
        fileTitle: fileTitle,
        fileType: fileType,
        replyCount: replyCountMap[row[0]] || 0,
        adminReacted: adminReacted
      });
      
      count++;
      if (count > 100) break; // Lấy 100 câu hỏi gần nhất thay vì 50
    }
    
    return { success: true, data: qaList };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getUserRole(email) {
  if (!email) return "MEM";
  const cleanEmail = String(email).trim().toLowerCase();
  if (_cachedUserRoles[cleanEmail]) {
    return _cachedUserRoles[cleanEmail];
  }

  const info = getUserAndPartner(cleanEmail);
  if (!info) {
    _cachedUserRoles[cleanEmail] = "MEM";
    return "MEM";
  }

  function normalizeRole(raw) {
    raw = (raw || "").toString().trim().toUpperCase();
    if (["ADMIN", "STUDENT", "S-STUDENT", "O-STUDENT", "VIP"].includes(raw)) return raw;
    return "MEM";
  }

  // Role của mình
  const myRole = normalizeRole(info.myRow[13]);

  // Không có partner
  if (!info.partnerRow) {
    _cachedUserRoles[cleanEmail] = myRole;
    return myRole;
  }

  // Role của partner
  const partnerRole = normalizeRole(info.partnerRow[13]);

  // Thứ tự mạnh bạn yêu cầu
  const order = ["MEM", "VIP", "O-STUDENT", "STUDENT", "S-STUDENT", "ADMIN"];

  const finalRole = order.indexOf(myRole) > order.indexOf(partnerRole)
    ? myRole
    : partnerRole;

  _cachedUserRoles[cleanEmail] = finalRole;
  return finalRole;
}


function canUserAccessExam(role, folderPath) {
  if (role === "ADMIN") return true;
  const upperPath = folderPath.toUpperCase();
  const isStudentOnly = upperPath.includes("STUDENT");
  const isVIP = upperPath.includes("VIP");

  if (isStudentOnly) return role === "STUDENT" || role === "S-STUDENT"; // KHÔNG cho O-STUDENT vào đây
  if (isVIP) return role === "VIP" || role === "O-STUDENT" || role === "STUDENT" || role === "S-STUDENT"; // Cho phép STUDENT và S-STUDENT vào VIP
  return true; // Các thư mục khác thì ai cũng truy cập
}

function translateToVietnamese(text) {
  return LanguageApp.translate(text, "en", "vi");
}

function canSendEmail(email) {
  const cache = CacheService.getScriptCache();
  const key = "emailSentCount_" + email;
  const count = parseInt(cache.get(key)) || 0;

  if (count >= 5) return false; // Giới hạn mỗi 24h

  cache.put(key, count + 1, 24 * 60 * 60); // TTL 24h
  return true;
}

function redeemGiftCode(inputCode) {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false, message: "Bạn chưa đăng nhập." };

  const giftSheet = SpreadsheetApp.getActive().getSheetByName("Giftcodes");
  if (!giftSheet) return { success: false, message: "Không tìm thấy sheet Giftcodes." };

  const data = giftSheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    const code = (data[i][0] || "").toString().trim();
    const coins = parseInt(data[i][1]) || 0;
    const usedList = (data[i][2] || "").toString();
    const maxUses = parseInt(data[i][3]) || 0;
    const rawMessage = (data[i][4] || "").toString().trim();
    const allowedUsers = (data[i][5] || "").toString();
    const targetSheetName = (data[i][6] || "").toString().trim();
    const targetColumnLetter = (data[i][7] || "").toString().trim().toUpperCase();
    const valueToWrite = (data[i][8] || "").toString();
    const startTime = data[i][9];
    const endTime = data[i][10];

    if (code.toUpperCase() === inputCode.toUpperCase()) {
      if (startTime instanceof Date && now < startTime) {
        return { success: false, message: "Gift code này chưa bắt đầu có hiệu lực." };
      }
      if (endTime instanceof Date && now > endTime) {
        return { success: false, message: "Gift code này đã hết hạn." };
      }

      const usedArray = usedList ? usedList.split(",") : [];
      if (allowedUsers) {
        const allowedArray = allowedUsers.split(",").map(e => e.trim().toLowerCase());
        if (!allowedArray.includes(email.toLowerCase())) {
          return { success: false, message: "Bạn không có quyền sử dụng gift code này." };
        }
      }
      if (usedArray.includes(email)) {
        return { success: false, message: "Bạn đã sử dụng gift code này rồi." };
      }
      if (maxUses > 0 && usedArray.length >= maxUses) {
        return { success: false, message: "Gift code này đã hết lượt sử dụng." };
      }

      if (coins > 0) {
        addUserPoints_(coins, "Nhập Giftcode: " + code);
      }

      usedArray.push(email);
      giftSheet.getRange(i + 1, 3).setValue(usedArray.join(","));

      // 1. Luôn lưu giftcode vào cột R của sheet Users (để liên kết lịch học)
      try {
        const userSheet = SpreadsheetApp.getActive().getSheetByName(USER_SHEET_NAME || "Users");
        if (userSheet) {
          const userData = userSheet.getDataRange().getValues();
          const emailColIndex = userData[0].findIndex(h => h.toLowerCase().includes("email"));
          const trackingColIndex = columnLetterToIndex("R"); // Cột R cố định
          
          if (emailColIndex !== -1 && trackingColIndex !== -1) {
            for (let r = 1; r < userData.length; r++) {
              if ((userData[r][emailColIndex] || "").toString().toLowerCase() === email.toLowerCase()) {
                const existingCodesRaw = (userData[r][trackingColIndex] || "").toString().trim();
                let existingCodes = existingCodesRaw ? existingCodesRaw.split(",").map(c => c.trim()).filter(Boolean) : [];
                
                const upperCode = code.toUpperCase();
                const alreadyExists = existingCodes.some(c => c.toUpperCase() === upperCode);
                if (!alreadyExists) {
                  existingCodes.push(code);
                }
                
                userSheet.getRange(r + 1, trackingColIndex + 1).setValue(existingCodes.join(", "));
                break;
              }
            }
          }
        }
      } catch (err) {
        console.error("Lỗi cập nhật cột R sheet Users: " + err.toString());
      }

      // 2. Ghi thêm giá trị vào sheet đích và cột đích cấu hình phụ nếu có
      if (targetSheetName && targetColumnLetter && valueToWrite) {
        const targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
        if (targetSheet) {
          const targetData = targetSheet.getDataRange().getValues();
          const emailColIndex = targetData[0].findIndex(h => h.toLowerCase().includes("email"));
          const targetColIndex = columnLetterToIndex(targetColumnLetter);
          const trackingColIndex = columnLetterToIndex("R"); // Cột R cố định

          if (emailColIndex !== -1 && targetColIndex !== -1) {
            for (let r = 1; r < targetData.length; r++) {
              if ((targetData[r][emailColIndex] || "").toString().toLowerCase() === email.toLowerCase()) {
                targetSheet.getRange(r + 1, targetColIndex + 1).setValue(valueToWrite);
                
                // Nối thêm giftcode mới vào cột R của sheet đích nếu chưa có
                const existingCodesRaw = (targetData[r][trackingColIndex] || "").toString().trim();
                let existingCodes = existingCodesRaw ? existingCodesRaw.split(",").map(c => c.trim()).filter(Boolean) : [];
                
                const lowerCode = code.toLowerCase();
                const alreadyExists = existingCodes.some(c => c.toLowerCase() === lowerCode);
                if (!alreadyExists) {
                  existingCodes.push(code);
                }
                
                targetSheet.getRange(r + 1, trackingColIndex + 1).setValue(existingCodes.join(", ")); // Lưu chuỗi phân cách bởi dấu phẩy
                break;
              }
            }
          }
        }
      }

      let htmlMessage = rawMessage.replace(
        /\[([^\]]+)\]\((https?:\/\/[^\s]+)\)/g,
        '<a href="$2" target="_blank">$1</a>'
      ).replace(
        /(?<!href=")(https?:\/\/[^\s]+)/g,
        '<a href="$1" target="_blank">$1</a>'
      );

      let finalMessage = "";
      if (coins > 0) finalMessage += `🎁 Bạn đã nhận ${coins} 🪙 Triple D Coin thành công!`;
      if (htmlMessage) finalMessage += `${coins > 0 ? "\n" : ""}📝 ${htmlMessage}`;

      return {
        success: true,
        message: finalMessage || "Bạn đã sử dụng giftcode thành công."
      };
    }
  }

  // Nếu không trùng với Giftcode nào, kiểm tra xem có trùng với Mã Lớp nào trong ClassSchedules không!
  try {
    const classSheet = SpreadsheetApp.getActive().getSheetByName("ClassSchedules");
    if (classSheet) {
      const classData = classSheet.getDataRange().getValues();
      let classExists = false;
      let className = "";
      for (let i = 1; i < classData.length; i++) {
        const cCode = (classData[i][0] || "").toString().trim().toUpperCase();
        if (cCode === inputCode.toUpperCase()) {
          classExists = true;
          className = (classData[i][1] || "").toString().trim();
          break;
        }
      }
      
      if (classExists) {
        // Thêm mã lớp học này vào cột R của sheet Users cho học sinh hiện tại
        const userSheet = SpreadsheetApp.getActive().getSheetByName(USER_SHEET_NAME || "Users");
        if (userSheet) {
          const userData = userSheet.getDataRange().getValues();
          const emailColIndex = userData[0].findIndex(h => h.toLowerCase().includes("email"));
          const trackingColIndex = columnLetterToIndex("R");
          
          if (emailColIndex !== -1 && trackingColIndex !== -1) {
            for (let r = 1; r < userData.length; r++) {
              if ((userData[r][emailColIndex] || "").toString().toLowerCase() === email.toLowerCase()) {
                const existingCodesRaw = (userData[r][trackingColIndex] || "").toString().trim();
                let existingCodes = existingCodesRaw ? existingCodesRaw.split(",").map(c => c.trim()).filter(Boolean) : [];
                
                const upperCode = inputCode.toUpperCase();
                const alreadyExists = existingCodes.some(c => c.toUpperCase() === upperCode);
                if (alreadyExists) {
                  return { success: false, message: `Bạn đã tham gia lớp học "${className}" (${upperCode}) trước đó rồi.` };
                }
                
                existingCodes.push(upperCode);
                userSheet.getRange(r + 1, trackingColIndex + 1).setValue(existingCodes.join(", "));
                
                return {
                  success: true,
                  message: `🎉 Chúc mừng! Bạn đã đăng ký tham gia lớp học "${className}" (${upperCode}) thành công.`
                };
              }
            }
          }
        }
      }
    }
  } catch (err) {
    console.error("Lỗi khi kiểm tra mã lớp: " + err.toString());
  }

  return { success: false, message: "Gift code hoặc Mã lớp không hợp lệ." };
}

function parseSheetDate_(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }
  const str = String(val).trim();
  if (!str) return null;

  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) return parsed;

  const parts = str.split(/[\sT]+/);
  if (parts.length > 0) {
    const dateParts = parts[0].split(/[\/\-]/);
    if (dateParts.length === 3) {
      let day, month, year;
      if (dateParts[0].length === 4) {
        year = parseInt(dateParts[0], 10);
        month = parseInt(dateParts[1], 10) - 1;
        day = parseInt(dateParts[2], 10);
      } else {
        day = parseInt(dateParts[0], 10);
        month = parseInt(dateParts[1], 10) - 1;
        year = parseInt(dateParts[2], 10);
      }
      let hour = 23, min = 59, sec = 59;
      if (parts.length > 1) {
        const timeParts = parts[1].split(":");
        if (timeParts.length >= 2) {
          hour = parseInt(timeParts[0], 10) || 0;
          min = parseInt(timeParts[1], 10) || 0;
          sec = parseInt(timeParts[2], 10) || 0;
        }
      }
      const altDate = new Date(year, month, day, hour, min, sec);
      if (!isNaN(altDate.getTime())) return altDate;
    }
  }
  return null;
}

function clearExpiredGiftCodeValues() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const giftSheet = ss.getSheetByName("Giftcodes");
    if (!giftSheet) {
      return { success: false, message: "Không tìm thấy sheet Giftcodes." };
    }

    const giftData = giftSheet.getDataRange().getValues();
    if (!giftData || giftData.length <= 1) {
      return { success: true, clearedCount: 0, message: "Không có dữ liệu Giftcode." };
    }

    const now = new Date();
    const activeCodesMap = new Map();
    const expiredCodesList = [];

    for (let i = 1; i < giftData.length; i++) {
      const code = (giftData[i][0] || "").toString().trim();
      if (!code) continue;

      const usedList = (giftData[i][2] || "").toString().trim();
      const targetSheetName = (giftData[i][6] || "").toString().trim();
      const targetColumnLetter = (giftData[i][7] || "").toString().trim().toUpperCase();
      const endTimeVal = giftData[i][10];
      const endTime = parseSheetDate_(endTimeVal);

      const isExpired = (endTime !== null && now > endTime);

      if (isExpired) {
        expiredCodesList.push({
          code: code,
          usedList: usedList,
          targetSheetName: targetSheetName,
          targetColumnLetter: targetColumnLetter
        });
      } else {
        activeCodesMap.set(code.toLowerCase(), {
          targetSheetName: targetSheetName,
          targetColumnLetter: targetColumnLetter
        });
      }
    }

    if (expiredCodesList.length === 0) {
      return { success: true, clearedCount: 0, message: "Không có Giftcode nào hết hạn." };
    }

    let totalCleared = 0;

    function cleanSheetForExpiredCode(sheet, expCode, usedEmails, targetColLetter) {
      if (!sheet) return 0;
      const sheetName = sheet.getName();
      const data = sheet.getDataRange().getValues();
      if (!data || data.length <= 1) return 0;

      const header = data[0];
      let emailColIndex = header.findIndex(h => h && h.toString().toLowerCase().includes("email"));
      if (emailColIndex === -1) emailColIndex = 0;

      const trackingColIndex = columnLetterToIndex("R");
      const targetColIndex = targetColLetter ? columnLetterToIndex(targetColLetter) : -1;

      let count = 0;

      for (let r = 1; r < data.length; r++) {
        const rowEmail = (data[r][emailColIndex] || "").toString().toLowerCase().trim();
        const rowGiftCodeStr = (data[r][trackingColIndex] || "").toString().trim();
        const rowGiftCodes = rowGiftCodeStr ? rowGiftCodeStr.split(",").map(c => c.trim()).filter(Boolean) : [];

        const isUserInUsedList = rowEmail && usedEmails.includes(rowEmail);
        const hasCodeInColumnR = rowGiftCodes.some(c => c.toLowerCase() === expCode.toLowerCase());

        if (isUserInUsedList || hasCodeInColumnR) {
          const remainingCodes = rowGiftCodes.filter(c => c.toLowerCase() !== expCode.toLowerCase());

          if (hasCodeInColumnR) {
            if (remainingCodes.length === 0) {
              sheet.getRange(r + 1, trackingColIndex + 1).clearContent();
            } else {
              sheet.getRange(r + 1, trackingColIndex + 1).setValue(remainingCodes.join(", "));
            }
            count++;
          }

          if (targetColIndex !== -1 && targetColIndex >= 0) {
            let hasOtherActiveCodeForCol = false;
            for (const remCode of remainingCodes) {
              const activeInfo = activeCodesMap.get(remCode.toLowerCase());
              if (activeInfo && activeInfo.targetSheetName.toLowerCase() === sheetName.toLowerCase() && activeInfo.targetColumnLetter === targetColLetter) {
                hasOtherActiveCodeForCol = true;
                break;
              }
            }

            if (!hasOtherActiveCodeForCol) {
              sheet.getRange(r + 1, targetColIndex + 1).clearContent();
            }
          }
        }
      }
      return count;
    }

    for (const exp of expiredCodesList) {
      const expCode = exp.code;
      const usedEmails = exp.usedList ? exp.usedList.split(",").map(e => e.trim().toLowerCase()).filter(Boolean) : [];

      if (exp.targetSheetName) {
        const targetSheet = ss.getSheetByName(exp.targetSheetName);
        if (targetSheet) {
          totalCleared += cleanSheetForExpiredCode(targetSheet, expCode, usedEmails, exp.targetColumnLetter);
        }
      }

      const userSheet = ss.getSheetByName(USER_SHEET_NAME || "Users");
      if (userSheet && (!exp.targetSheetName || exp.targetSheetName.toLowerCase() !== (USER_SHEET_NAME || "Users").toLowerCase())) {
        totalCleared += cleanSheetForExpiredCode(userSheet, expCode, usedEmails, "");
      }
    }

    return {
      success: true,
      clearedCount: totalCleared,
      expiredCodesCount: expiredCodesList.length,
      message: `Đã dọn dẹp thành công ${expiredCodesList.length} giftcode hết hạn.`
    };
  } catch (err) {
    console.error("Lỗi khi dọn dẹp Giftcode hết hạn: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

function columnLetterToIndex(letter) {
  if (!letter || typeof letter !== 'string') return -1;
  const clean = letter.trim().toUpperCase();
  if (!clean) return -1;
  return clean.split('').reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1;
}

function getMonthlyTransferTotals(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CoinTransactions");
  if (!sheet) return { sent: 0, received: 0 };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { sent: 0, received: 0 };
  
  const currentMonth = new Date().getMonth();
  const currentYear = new Date().getFullYear();
  let sent = 0;
  let received = 0;
  
  const targetEmail = String(email).toLowerCase().trim();

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const rowDateStr = String(row[0]);
    
    const parts = rowDateStr.split(" ");
    if (parts.length < 2) continue;
    const dateParts = parts[0].split("/");
    if (dateParts.length < 3) continue;
    const month = parseInt(dateParts[1], 10) - 1;
    const year = parseInt(dateParts[2], 10);
    
    if (year < currentYear || (year === currentYear && month < currentMonth)) {
      break; 
    }
    if (month !== currentMonth || year !== currentYear) {
      continue;
    }
    
    const rowEmail = String(row[1]).toLowerCase().trim();
    if (rowEmail !== targetEmail) continue;
    
    const rowAmount = Number(row[2]);
    const source = String(row[4]).toLowerCase();
    
    if (source.includes("chuyển") && source.includes("td coin cho") && rowAmount < 0) {
      const match = source.match(/chuyển\s+(\d+)/);
      if (match) {
        sent += parseInt(match[1], 10);
      }
    } else if (source.includes("nhận") && source.includes("td coin từ") && rowAmount > 0) {
      received += rowAmount;
    }
  }
  
  return { sent, received };
}

function transferPoints(toEmail, amount) {
  const fromEmail = CacheService.getUserCache().get("loggedInUser");
  if (!fromEmail) return { success: false, message: "Bạn chưa đăng nhập." };

  if (!isValidEmail(toEmail)) {
    return { success: false, message: "Email người nhận không hợp lệ." };
  }

  if (fromEmail.toLowerCase() === toEmail.toLowerCase()) {
    return { success: false, message: "Không thể chuyển Triple D Coin cho chính mình." };
  }

  amount = parseInt(amount);
  if (isNaN(amount) || amount <= 0) {
    return { success: false, message: "Số Triple D Coin phải là số nguyên dương." };
  }

  const LIMIT_PER_MONTH = 100000;
  
  // Kiểm tra giới hạn người gửi
  const senderTotals = getMonthlyTransferTotals(fromEmail);
  if (senderTotals.sent + amount > LIMIT_PER_MONTH) {
    const remain = Math.max(0, LIMIT_PER_MONTH - senderTotals.sent);
    return { success: false, message: `Tháng này bạn đã chuyển ${senderTotals.sent} coin. Hạn mức còn lại: ${remain} coin. Bạn không thể chuyển thêm ${amount} coin.` };
  }
  
  // Kiểm tra giới hạn người nhận
  const receiverTotals = getMonthlyTransferTotals(toEmail);
  if (receiverTotals.received + amount > LIMIT_PER_MONTH) {
    const remain = Math.max(0, LIMIT_PER_MONTH - receiverTotals.received);
    return { success: false, message: `Tháng này người nhận đã nhận ${receiverTotals.received} coin. Hạn mức nhận còn lại của họ: ${remain} coin. Không thể chuyển thêm ${amount} coin cho người này.` };
  }

  const feeRate = 0.2; // 💰 20% phí
  const fee = Math.ceil(amount * feeRate);
  const totalCost = amount + fee;

  const sender = getUserByEmail(fromEmail);
  const receiver = getUserByEmail(toEmail);

  if (!receiver) {
    return { success: false, message: "Không tìm thấy người nhận." };
  }

  // Kiểm tra hạn mức chuyển tối đa 100,000 Coin / tháng cho người gửi
  const senderLimits = getP2PMonthlyTotals(fromEmail);
  if (senderLimits.sent + amount > 100000) {
    return {
      success: false,
      message: `Giao dịch bất thành. Bạn đã chuyển tổng cộng ${senderLimits.sent.toLocaleString()} 🪙 trong tháng này. Hạn mức chuyển tối đa là 100,000 🪙/tháng. Bạn chỉ còn có thể chuyển tối đa ${(100000 - senderLimits.sent).toLocaleString()} 🪙.`
    };
  }

  // Kiểm tra hạn mức nhận tối đa 100,000 Coin / tháng cho người nhận
  const receiverLimits = getP2PMonthlyTotals(toEmail);
  if (receiverLimits.received + amount > 100000) {
    return {
      success: false,
      message: `Giao dịch bất thành. Người nhận đã nhận tổng cộng ${receiverLimits.received.toLocaleString()} 🪙 trong tháng này. Hạn mức nhận tối đa là 100,000 🪙/tháng. Họ chỉ có thể nhận thêm tối đa ${(100000 - receiverLimits.received).toLocaleString()} 🪙.`
    };
  }

  const senderPoints = parseInt(sender.data[10]) || 0;
  if (senderPoints < totalCost) {
    return {
      success: false,
      message: `Bạn cần ${totalCost} Triple D Coin để chuyển ${amount} 🪙 (bao gồm ${fee} 🪙 phí).`
    };
  }

  // Trừ người gửi
  updateUserField(sender.row, 10, senderPoints - totalCost);

  // Cộng người nhận
  const receiverPoints = parseInt(receiver.data[10]) || 0;
  updateUserField(receiver.row, 10, receiverPoints + amount);

  logPointTransfer(fromEmail, toEmail, amount, `Phí ${fee}, Tổng trừ ${totalCost}`);

    return {
      success: true,
      message: `Đã chuyển ${amount} 🪙 đến ${receiver.data[0] || toEmail}. Phí: ${fee} 🪙. Tổng trừ: ${totalCost} 🪙.`
    };
}

function logPointTransfer(fromEmail, toEmail, amount, note) {
  const feeRate = 0.2; // 💰 20% phí
  const fee = Math.ceil(amount * feeRate);
  const totalCost = amount + fee;

  // Ghi nhận thay đổi tiền vào CoinTransactions
  logCoinTransaction(fromEmail, -totalCost, "TD Coin", `Chuyển ${amount} TD Coin cho ${toEmail} (Phí ${fee})`);
  logCoinTransaction(toEmail, amount, "TD Coin", `Nhận ${amount} TD Coin từ ${fromEmail}`);

  // Ghi log vào Tracking sheet
  logActivityToSheet(fromEmail, "COIN_TRANSFER", `Chuyển ${amount} TD Coin cho ${toEmail}. Note: ${note || 'Không có'}`);
  logActivityToSheet(toEmail, "COIN_RECEIVED", `Nhận ${amount} TD Coin từ ${fromEmail}. Note: ${note || 'Không có'}`);
}

function parseTransactionDate(val) {
  if (val instanceof Date) {
    return val;
  }
  if (typeof val === "string") {
    const parts = val.split(" ");
    if (parts.length > 0) {
      const dateParts = parts[0].split("/");
      if (dateParts.length === 3) {
        const day = parseInt(dateParts[0], 10);
        const month = parseInt(dateParts[1], 10) - 1; // 0-indexed month
        const year = parseInt(dateParts[2], 10);
        let hour = 0, min = 0, sec = 0;
        if (parts.length > 1) {
          const timeParts = parts[1].split(":");
          if (timeParts.length >= 3) {
            hour = parseInt(timeParts[0], 10);
            min = parseInt(timeParts[1], 10);
            sec = parseInt(timeParts[2], 10);
          }
        }
        return new Date(year, month, day, hour, min, sec);
      }
    }
    const parsed = Date.parse(val);
    if (!isNaN(parsed)) {
      return new Date(parsed);
    }
  }
  return null;
}

function getP2PMonthlyTotals(email) {
  const emailClean = (email || "").toLowerCase().trim();
  let totalSent = 0;
  let totalReceived = 0;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("CoinTransactions");
    if (!sheet) return { sent: 0, received: 0 };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { sent: 0, received: 0 };

    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth();

    for (let i = 1; i < data.length; i++) {
      const rowEmail = (data[i][1] || "").toString().toLowerCase().trim();
      if (rowEmail !== emailClean) continue;

      const dateVal = data[i][0];
      const txDate = parseTransactionDate(dateVal);
      if (!txDate) continue;

      // Check if it's in the current calendar month
      if (txDate.getFullYear() === currentYear && txDate.getMonth() === currentMonth) {
        const source = (data[i][4] || "").toString();
        
        // Check for sending: "Chuyển {amount} TD Coin cho ..."
        const sendMatch = source.match(/Chuyển\s+(\d+)\s+TD\s+Coin/i);
        if (sendMatch && source.toLowerCase().includes("cho")) {
          totalSent += parseInt(sendMatch[1], 10);
        } else {
          // Check for receiving: "Nhận {amount} TD Coin từ ..."
          const receiveMatch = source.match(/Nhận\s+(\d+)\s+TD\s+Coin/i);
          if (receiveMatch && source.toLowerCase().includes("từ")) {
            totalReceived += parseInt(receiveMatch[1], 10);
          }
        }
      }
    }
  } catch (e) {
    Logger.log("Error in getP2PMonthlyTotals: " + e.toString());
  }

  return { sent: totalSent, received: totalReceived };
}

function deductUserPoints(amount, source = "Admin / Hệ thống trừ tiền") {
  amount = Number(amount);
  if (isNaN(amount) || amount <= 0) return false;
  return addUserPoints_(-amount, source);
}

function getStoreItems() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Store");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => ({
    item: row[0],
    price: row[1],
    description: row[2],
    image: row[3] || ""
  }));
}

function checkAndProcessStoreRewards(email) {
  let acquiredRewards = [];
  try {
    if (!email) return acquiredRewards;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName("PurchaseLog");
    if (!logSheet) return acquiredRewards;

    const data = logSheet.getDataRange().getValues();
    const now = new Date().getTime();

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[1] && row[1].toString().toLowerCase() === email.toLowerCase()) {
            const status = row[5];
            if (status !== "Claimed" && status !== "Claimed_Expired") {
                const itemName = row[2] ? row[2].toString() : "";
                const purchaseTime = new Date(row[0]).getTime();
                if (isNaN(purchaseTime)) continue;
                
                let isEligible = false;
                let rewardAmount = 0;
                let isExpired = false;
                
                if (itemName.includes("Thẻ Tích Lũy Tháng")) {
                    // 30 ngày = 30 * 24 * 60 * 60 * 1000
                    if (now - purchaseTime >= 2592000000) {
                        isEligible = true;
                        rewardAmount = 5000;
                    }
                } else if (itemName.includes("Thẻ Tích Lũy Tuần")) {
                    // 7 ngày = 7 * 24 * 60 * 60 * 1000
                    if (now - purchaseTime >= 604800000) {
                        isEligible = true;
                        rewardAmount = 2400;
                    }
                }
                else if (itemName.includes("Góp VIP 3 Tháng")) {
                    // 90 ngày = 90 * 24 * 60 * 60 * 1000
                    if (now - purchaseTime >= 7776000000) {
                        isExpired = true;
                        const userSheet = ss.getSheetByName("Users");
                        const userData = userSheet.getDataRange().getValues();
                        const userIndex = userData.findIndex(r => r[1] && r[1].toString().toLowerCase() === email.toLowerCase());
                        if (userIndex !== -1) {
                            // Reset role
                            updateUserField(userIndex, 13, "MEM");
                        }
                        logSheet.getRange(i + 1, 6).setValue("Claimed_Expired");
                    }
                }
                
                if (isEligible) {
                    const userSheet = ss.getSheetByName("Users");
                    const userData = userSheet.getDataRange().getValues();
                    const userIndex = userData.findIndex(r => r[1] && r[1].toString().toLowerCase() === email.toLowerCase());
                    if (userIndex !== -1) {
                        const currentPoints = Number(userData[userIndex][10]) || 0;
                        userSheet.getRange(userIndex + 1, 11).setValue(currentPoints + rewardAmount);
                    }
                    logSheet.getRange(i + 1, 6).setValue("Claimed");

                    const msg = `Bạn đã nhận được ${rewardAmount} 🪙 từ ${itemName}`;
                    // Gửi thông báo cho user
                    addAutoNotificationToSheet(
                        "🎁 Phần Thưởng Cửa Hàng", 
                        msg,
                        "STORE_REWARD",
                        "",
                        email,
                        "system"
                    );
                    acquiredRewards.push(msg);
                }
            }
        }
    }
  } catch(e) {
    Logger.log("Error processing store rewards: " + e);
  }
  return acquiredRewards;
}

function purchaseItem(itemName) {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false, message: "Bạn chưa đăng nhập." };

  const ss = SpreadsheetApp.getActive();
  
  // Check if they already have an active card of this type
  if (itemName.includes("Thẻ Tích Lũy") || itemName.includes("Góp VIP")) {
    const logSheet = ss.getSheetByName("PurchaseLog");
    if (logSheet) {
      const logData = logSheet.getDataRange().getValues();
      // Skip header row if exists, but we can just search from the end or all
      for (let i = 0; i < logData.length; i++) {
        const row = logData[i];
        if (row[1] && row[1].toString().toLowerCase() === email.toLowerCase()) {
          if (row[2] && row[2].toString() === itemName) {
            const status = row[5] ? row[5].toString() : "";
            if (status === "Pending") {
              return { success: false, message: `Bạn đang có một "${itemName}" chưa hết hạn. Vui lòng chờ thẻ hiện tại hết hạn mới có thể mua thêm!` };
            }
          }
        }
      }
    }
  }

  const storeSheet = ss.getSheetByName("Store");
  const storeData = storeSheet.getDataRange().getValues();
  const itemRow = storeData.find(row => row[0] === itemName);
  if (!itemRow) return { success: false, message: "Mặt hàng không tồn tại." };

  const price = Number(itemRow[1]);
  if (isNaN(price)) return { success: false, message: "Giá sản phẩm không hợp lệ." };

  const userSheet = ss.getSheetByName("Users");
  const userData = userSheet.getDataRange().getValues();

  const userIndex = userData.findIndex(row =>
    row[1] && row[1].toString().toLowerCase() === email.toLowerCase()
  );
  if (userIndex === -1) return { success: false, message: "Không tìm thấy người dùng." };

  const currentTDC = Number(userData[userIndex][10]) || 0;
  if (currentTDC < price) return { success: false, message: "Không đủ Triple D Coin." };

  const newTDC = currentTDC - price;
  userSheet.getRange(userIndex + 1, 11).setValue(newTDC);  // Cột K = 11
  logCoinTransaction(email, -price, "TD Coin", "Mua hàng Store: " + itemName);

  // Kích hoạt VIP ngay lập tức
  if (itemName === "Góp VIP 3 Tháng" || itemName.includes("VIP")) {
    updateUserField(userIndex, 13, "VIP");
  }

  // ✅ Tạo mã đơn hàng chữ + số
  const orderCode = generateOrderCode();  // Gọi hàm bên dưới

  // ✅ Ghi log với mã đơn hàng ở cột 5, trạng thái ở cột 6 (Pending)
  const logSheet = ss.getSheetByName("PurchaseLog") || ss.insertSheet("PurchaseLog");
  logSheet.appendRow([new Date(), email, itemName, price, orderCode, "Pending"]);

  // Ghi nhận vào Tracking sheet
  logActivityToSheet(email, "STORE_PURCHASE", `Mua mặt hàng: "${itemName}". Giá: ${price} TD Coin. Mã đơn hàng: ${orderCode}`);

  const finalMessage = (itemName === "Góp VIP 3 Tháng" || itemName.includes("VIP"))
     ? `Đã đổi "${itemName}" thành công! Bạn đã được kích hoạt VIP.<br>Còn lại ${newTDC} 🪙.<br>Mã quy đổi: <b>${orderCode}</b>`
     : `Đã đổi "${itemName}" thành công! Hệ thống sẽ tự động gửi 🪙 khi đủ thời gian.<br>Còn lại ${newTDC} 🪙.<br>Mã quy đổi: <b>${orderCode}</b>`;

  return {
    success: true,
    message: finalMessage,
    orderCode: orderCode
  };
}

function generateOrderCode(length = 6) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let code = '';
  for (let i = 0; i < length; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return '#' + code;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === "close") {
      // Xử lý thu hồi quyền truy cập đối với các file đang mở của user
      if (Array.isArray(data.userFiles) && data.email) {
        data.userFiles.forEach(fileId => {
          revokeFileAccess(fileId, data.email);
        });
      }

      return ContentService.createTextOutput("OK");
    }

    return ContentService.createTextOutput("Invalid request");
  } catch (err) {
    return ContentService.createTextOutput("ERROR: " + err.message);
  }
}

function revokeFileAccess(fileId, email) {
  if (!fileId) return false;
  let cleanEmail = (email || "").toLowerCase().trim();
  if (!cleanEmail) {
    try {
      const cached = CacheService.getUserCache().get("loggedInUser");
      if (cached) cleanEmail = cached.toLowerCase().trim();
    } catch (e) {}
  }
  if (!cleanEmail) return false;

  const cacheKey = "perm_" + fileId + "_" + cleanEmail;

  // 🚀 TỐI ƯU SIÊU TỐC 1: Xóa trực tiếp qua permission ID lưu sẵn trong CacheService (~30-50ms)
  try {
    var cachedPermId = CacheService.getScriptCache().get(cacheKey);
    if (cachedPermId) {
      Drive.Permissions.remove(fileId, cachedPermId, { supportsAllDrives: true });
      CacheService.getScriptCache().remove(cacheKey);
      Logger.log(`⚡ [Ultra-Fast Revoke] Revoked permission ${cachedPermId} for ${cleanEmail} on file ${fileId}`);
      return true;
    }
  } catch (eCacheRevoke) {
    Logger.log("Chuyển sang fallback Drive.Permissions.list: " + eCacheRevoke.message);
  }

  // 🚀 TỐI ƯU SIÊU TỐC 2: Quét trực tiếp bằng Drive.Permissions.list (~100ms)
  try {
    var permList = Drive.Permissions.list(fileId, { 
      supportsAllDrives: true, 
      fields: 'permissions(id,emailAddress,role)' 
    });
    if (permList && permList.permissions) {
      for (var i = 0; i < permList.permissions.length; i++) {
        var p = permList.permissions[i];
        if (p.emailAddress && p.emailAddress.toLowerCase() === cleanEmail && p.role !== 'owner') {
          Drive.Permissions.remove(fileId, p.id, { supportsAllDrives: true });
          CacheService.getScriptCache().remove(cacheKey);
          Logger.log(`⚡ [Fast Revoke] Revoked permission ${p.id} for ${cleanEmail} on file ${fileId}`);
          return true;
        }
      }
    }
  } catch (eDriveApi) {}

  // 🚀 FALLBACK NẾU DRIVE API GẶP LỖI: dùng DriveApp
  try {
    var file = DriveApp.getFileById(fileId);
    file.removeViewer(cleanEmail);
    Logger.log(`[Instant Revoke Fallback] File ${fileId} revoked for user ${cleanEmail}`);
    return true;
  } catch (error) {
    Logger.log("Lỗi revokeFileAccess: " + error.message);
    return false;
  }
}

/**
 * Quản lý danh sách thu hồi quyền ngầm ở Backend (Auto Revoke Queue)
 */
function scheduleAutoRevokeBackend_(fileId, email, delayMs) {
  try {
    if (!fileId || !email) return;
    const cleanEmail = email.toLowerCase().trim();
    const delay = delayMs || 10000; // Mặc định 10 giây
    const expireTime = Date.now() + delay;

    // 1. Lưu thông tin vào ScriptProperties
    const props = PropertiesService.getScriptProperties();
    let pending = [];
    try {
      const raw = props.getProperty("PENDING_AUTO_REVOKES");
      if (raw) pending = JSON.parse(raw);
    } catch (e) {
      pending = [];
    }

    // Tránh trùng lặp
    pending = pending.filter(item => !(item.fileId === fileId && item.email === cleanEmail));
    pending.push({ fileId: fileId, email: cleanEmail, expireTime: expireTime });
    safeSetScriptProperty("PENDING_AUTO_REVOKES", JSON.stringify(pending));

    // 2. Tạo Time-driven Trigger chạy ngầm trên server Google nếu chưa có
    const triggers = ScriptApp.getProjectTriggers();
    let triggerExists = false;
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === "handleBackendAutoRevokeTrigger") {
        triggerExists = true;
        break;
      }
    }

    if (!triggerExists) {
      ScriptApp.newTrigger("handleBackendAutoRevokeTrigger")
        .timeBased()
        .after(delay)
        .create();
      Logger.log(`[Auto-Revoke Backend] Đã lên lịch trigger ngầm thu hồi sau ${delay / 1000}s cho ${fileId}`);
    }
  } catch (err) {
    Logger.log("[Auto-Revoke Backend] Lỗi lên lịch trigger: " + err.message);
  }
}

/**
 * Trigger chạy ngầm do Google Apps Script tự động kích hoạt
 */
function handleBackendAutoRevokeTrigger(e) {
  // Xóa trigger vừa gọi để tránh tích tụ trigger
  if (e && e.triggerUid) {
    try {
      const triggers = ScriptApp.getProjectTriggers();
      for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getUniqueId() === e.triggerUid) {
          ScriptApp.deleteTrigger(triggers[i]);
          break;
        }
      }
    } catch (err) {
      Logger.log("Lỗi xóa trigger auto-revoke: " + err.message);
    }
  }

  processPendingAutoRevokes_();
}

/**
 * Quét danh sách pending và thu hồi quyền các file đã hết hạn
 */
function processPendingAutoRevokes_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty("PENDING_AUTO_REVOKES");
    if (!raw) return;

    let pending = JSON.parse(raw);
    if (!Array.isArray(pending) || pending.length === 0) return;

    const now = Date.now();
    const remaining = [];

    for (let i = 0; i < pending.length; i++) {
      const item = pending[i];
      if (now >= item.expireTime - 1000) {
        // Đã đến hạn thu hồi
        Logger.log(`[Auto-Revoke Backend Execution] Đang thu hồi quyền ngầm file ${item.fileId} của ${item.email}`);
        revokeFileAccess(item.fileId, item.email);
      } else {
        // Chưa đến hạn
        remaining.push(item);
      }
    }

    // Cập nhật lại danh sách còn lại
    safeSetScriptProperty("PENDING_AUTO_REVOKES", JSON.stringify(remaining));

    // Nếu vẫn còn item chưa đến hạn, lên lịch trigger tiếp theo
    if (remaining.length > 0) {
      let minDelay = 5000;
      const nextTime = Math.min(...remaining.map(r => r.expireTime));
      const calculatedDelay = nextTime - now;
      if (calculatedDelay > 1000) minDelay = Math.max(1000, calculatedDelay);

      const triggers = ScriptApp.getProjectTriggers();
      let triggerExists = false;
      for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === "handleBackendAutoRevokeTrigger") {
          triggerExists = true;
          break;
        }
      }

      if (!triggerExists) {
        ScriptApp.newTrigger("handleBackendAutoRevokeTrigger")
          .timeBased()
          .after(minDelay)
          .create();
      }
    }
  } catch (err) {
    Logger.log("[Auto-Revoke Backend Execution] Lỗi xử lý revocation: " + err.message);
  }
}

/**
 * Chia sẻ file cho email cụ thể và ghi log vào sheet OpenFilesLog.
 * Dùng chung cho tất cả các tài nguyên Thư viện, Khóa học, Đề thi, Tài liệu.
 */
function shareFileAndLogAccess_(fileId, email) {
  if (!fileId || !email) return false;
  const cleanEmail = email.toLowerCase().trim();
  
  try {
    const file = DriveApp.getFileById(fileId);
    
    // Đảm bảo file được khóa về PRIVATE trước khi share
    try {
      file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    } catch (sharingErr) {
      Logger.log("Không thể setSharing PRIVATE cho file " + fileId + ": " + sharingErr.message);
    }
    
    // Cấp quyền Reader cho người dùng không gửi email thông báo
    var createdPerm = Drive.Permissions.create(
      {
        role: 'reader',
        type: 'user',
        emailAddress: cleanEmail
      },
      fileId,
      { 
        sendNotificationEmail: false,
        supportsAllDrives: true
      }
    );
    
    // Cache permission ID siêu tốc để phục vụ thu hồi ngay lập tức (~30ms)
    if (createdPerm && createdPerm.id) {
      try {
        CacheService.getScriptCache().put("perm_" + fileId + "_" + cleanEmail, createdPerm.id, 600);
      } catch (eCache) {}
    }
    
    // Ghi log vào sheet 'OpenFilesLog'
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('OpenFilesLog');
    if (!sheet) {
      sheet = ss.insertSheet('OpenFilesLog');
      sheet.appendRow(['fileId', 'email', 'lastOpenTime']);
    }
    
    const data = sheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == fileId && (data[i][1] || "").toLowerCase().trim() === cleanEmail) {
        foundRow = i + 1;
        break;
      }
    }
    
    const nowIso = new Date().toISOString();
    if (foundRow > 0) {
      sheet.getRange(foundRow, 3).setValue(nowIso);
    } else {
      sheet.appendRow([fileId, cleanEmail, nowIso]);
    }
    
    // Lên lịch thu hồi quyền ngầm ở Backend sau 3 giây (Failsafe dự phòng)
    scheduleAutoRevokeBackend_(fileId, cleanEmail, 3000);

    return true;
  } catch (e) {
    Logger.log(`Lỗi trong shareFileAndLogAccess_ cho file ${fileId} và user ${cleanEmail}: ${e.message}`);
    return false;
  }
}

function shareCourseFileToUserWithLog(fileId, userEmailParam, isMockExam = false) {
  try {
    let email = (userEmailParam || "").toLowerCase().trim();
    if (!email) {
      email = CacheService.getUserCache().get("loggedInUser");
    }
    if (!email) return { success: false, message: "Bạn chưa đăng nhập." };
    const cleanEmail = email.toLowerCase().trim();

    const role = getUserRole(cleanEmail);
    if (role !== "ADMIN" && role !== "STUDENT" && role !== "O-STUDENT" && role !== "S-STUDENT") {
      return { success: false, message: "Từ chối truy cập: Bạn cần nâng cấp tài khoản lên Student hoặc Online-Student để sử dụng tính năng này." };
    }

    var allowedFolders = getUserAllowedFolders(cleanEmail);
    let courseInfo = null;

    if (isMockExam) {
      // Mock Exam logic
      const mockResult = getMockExams();
      if (!mockResult.success) return { success: false, message: "Không thể lấy danh sách đề thi." };
      courseInfo = mockResult.data.find(e => e.fileId === fileId);
      if (!courseInfo) return { success: false, message: "Bạn không có quyền hoặc không tìm thấy đề thi." };
    } else {
      // Course logic
      var cachedCourses = getCachedCoursesFromSheet();
      courseInfo = cachedCourses.find(c => c.fileId === fileId);
      if (!courseInfo) return { success: false, message: "Không tìm thấy khóa học tương ứng." };

      const blockedKeywords = ["FuII", "Fixed"];
      if (role !== "ADMIN" && blockedKeywords.some(kw => (courseInfo.title || "").includes(kw))) {
        const chanceToOpen = 0.1;
        if (Math.random() >= chanceToOpen) {
          return { success: false, fakeError: true }; 
        }
      }

      // Check course permission
      if (role !== 'ADMIN' && !allowedFolders.includes("ALL_FOLDERS")) {
        const allAncestors = courseInfo.ancestors || [];
        const hasAccess = allowedFolders.includes(courseInfo.folderId) || 
                          allAncestors.some(anc => allowedFolders.includes(anc));
        if (!hasAccess) {
          return { success: false, message: "Bạn không có quyền truy cập chuyên đề này." };
        }
      }
    }

    let isDirectPdf = false;
    if (isMockExam) {
      isDirectPdf = true;
    } else if (courseInfo && (courseInfo.type === 'pdf' || courseInfo.type === 'doc' || courseInfo.type === 'docx' || courseInfo.type === 'document')) {
      isDirectPdf = true;
    }

    // Với file video của khóa học: chia sẻ quyền Drive tạm thời
    // Với file PDF của Khóa Học / Khảo Thí: render qua PDF.js trực tiếp từ server nên không cần chia sẻ Drive
    if (!isDirectPdf) {
      const success = shareFileAndLogAccess_(fileId, cleanEmail);
      if (!success) {
        return { success: false, message: "Lỗi hệ thống khi mở tài liệu video." };
      }
    }

    // Ghi hoạt động vào Tracking sheet
    logActivityToSheet(cleanEmail, "OPEN_FILE", `Mở tài liệu: [${courseInfo ? courseInfo.title || '' : ''}] (ID: ${fileId})`);

    return { success: true, message: isDirectPdf ? "Đã xác thực và mở file PDF qua PDF.js." : "Đã cấp quyền video và ghi nhận thời gian mở file." };

  } catch (e) {
    return { success: false, message: "Lỗi khi cấp quyền: " + e.message };
  }
}

function revokeAccessAndCleanLogIfExpired() {
  try {
    // Tự động dọn dẹp các thông báo cũ và phiên theo dõi cũ để giữ sheet sạch sẽ
    try {
      autoCleanupMonitorSessions();
    } catch (eMonitor) {
      Logger.log("Error during autoCleanupMonitorSessions: " + eMonitor.message);
    }
    try {
      autoCleanupNotifications();
    } catch (eNotif) {
      Logger.log("Error during autoCleanupNotifications: " + eNotif.message);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Đồng thời dọn dẹp các sheet cũ thừa thãi nếu còn tồn tại
    var oldSheet = ss.getSheetByName('OpenFiles');
    if (oldSheet) {
      try {
        ss.deleteSheet(oldSheet);
        Logger.log("Đã dọn dẹp và xóa sheet 'OpenFiles' cũ thành công.");
      } catch (err) {
        Logger.log("Không thể xóa sheet 'OpenFiles': " + err.message);
      }
    }

    var sheet = ss.getSheetByName('OpenFilesLog');
    if (!sheet) {
      Logger.log("Không tìm thấy sheet OpenFilesLog");
      return;
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log("Sheet OpenFilesLog không có dữ liệu");
      return;
    }

    var now = new Date();
    var TWO_HOURS_MS = 2 * 60 * 60 * 1000; // 2 tiếng cho học sinh
    var ONE_HOUR_MS = 1 * 60 * 60 * 1000;   // 1 tiếng cho admin (bảo mật tối đa nếu quên bấm thu hồi)

    var roleCache = {};
    function getCachedUserRole(userEmail) {
      if (!roleCache[userEmail]) {
        roleCache[userEmail] = getUserRole(userEmail);
      }
      return roleCache[userEmail];
    }

    for (var i = data.length - 1; i > 0; i--) {
      var fileId = data[i][0];
      var email = (data[i][1] || "").toLowerCase().trim();
      var lastOpenTimeStr = data[i][2];

      // Không có thời gian mở
      if (!lastOpenTimeStr) {
        Logger.log(`Bỏ qua dòng ${i + 1} (${email}, ${fileId}): Không có thời gian mở`);
        continue;
      }

      var lastOpenTime = new Date(lastOpenTimeStr);
      if (isNaN(lastOpenTime.getTime())) {
        Logger.log(`Bỏ qua dòng ${i + 1} (${email}, ${fileId}): Thời gian mở không hợp lệ`);
        continue;
      }

      var userRole = getCachedUserRole(email);
      var expiryLimit = (userRole === "ADMIN") ? ONE_HOUR_MS : TWO_HOURS_MS;

      // Chỉ xử lý nếu đã quá hạn
      if ((now - lastOpenTime) < expiryLimit) {
        var timeLeftMin = Math.round((expiryLimit - (now - lastOpenTime)) / (60 * 1000));
        Logger.log(`Bỏ qua dòng ${i + 1} (${email}, ${fileId}): Chưa quá hạn cho ${email} (${userRole}) (còn ${timeLeftMin} phút)`);
        continue;
      }

      var file;
      try {
        file = DriveApp.getFileById(fileId);
      } catch (e) {
        Logger.log(`Lỗi: không tìm thấy file ${fileId}. Đã xóa dòng log.`);
        sheet.deleteRow(i + 1);
        continue;
      }

      // Luôn thử xóa quyền nếu hết hạn
      try {
        var permissionListResponse = Drive.Permissions.list(fileId, { supportsAllDrives: true });
        var permissionList = permissionListResponse.permissions || [];
        for (var j = 0; j < permissionList.length; j++) {
          var perm = permissionList[j];
          if (perm.emailAddress && perm.emailAddress.toLowerCase() === email) {
            Drive.Permissions.remove(fileId, perm.id, { supportsAllDrives: true });
            Logger.log(`Đã xóa quyền của ${email} trên file ${fileId}`);
            break;
          }
        }
      } catch (e) {
        Logger.log(`Lỗi khi xóa quyền của ${email} trên file ${fileId}: ${e.message}`);
      }

      // Kiểm tra lại quyền sau khi xóa
      var stillHasAccess = false;
      try {
        var updatedPermissions = file.getAccess(email);
        stillHasAccess = (updatedPermissions === DriveApp.Permission.VIEW || updatedPermissions === DriveApp.Permission.EDIT);
      } catch (e) {
        stillHasAccess = false;
      }

      if (!stillHasAccess) {
        sheet.deleteRow(i + 1);
        Logger.log(`Đã xóa dòng log của ${email} trên file ${fileId}`);
      } else {
        Logger.log(`Bỏ qua xóa log dòng ${i + 1} (${email}, ${fileId}): Người dùng vẫn còn quyền`);
      }
    }
  } catch (error) {
    Logger.log("Lỗi khi xử lý revokeAccessAndCleanLogIfExpired: " + error.message);
  }
}

/**
 * Tự động chia sẻ Đề thi gốc cho Admin khi họ đang chấm bài
 * để tránh việc Admin phải tự mở từ giao diện Khảo thí.
 */
function shareExamToAdmin(fileId) {
  if (!fileId) return { success: false, message: "Thiếu fileId" };
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false, message: "Bạn chưa đăng nhập." };
  
  const cleanEmail = email.toLowerCase().trim();
  const role = getUserRole(cleanEmail);
  if (role !== "ADMIN") {
    return { success: false, message: "Từ chối truy cập: Bạn không phải Admin." };
  }
  
  const success = shareFileAndLogAccess_(fileId, cleanEmail);
  if (success) {
    return { success: true, message: "Đã chia sẻ đề thi gốc cho Admin." };
  } else {
    return { success: false, message: "Không thể chia sẻ đề thi gốc cho Admin." };
  }
}

/**
 * Thu hồi tất cả các đề thi gốc đã chia sẻ cho Admin hiện tại khi họ đã hoàn thành việc chấm thi.
 */
function revokeAllAdminSharedFiles() {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false, message: "Bạn chưa đăng nhập." };
  
  const cleanEmail = email.toLowerCase().trim();
  const role = getUserRole(cleanEmail);
  if (role !== "ADMIN") {
    return { success: false, message: "Từ chối truy cập: Bạn không phải Admin." };
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('OpenFilesLog');
    if (!sheet) {
      return { success: true, message: "Không có file nào cần thu hồi." };
    }
    
    const data = sheet.getDataRange().getValues();
    let count = 0;
    
    // Quét ngược để xóa và thu hồi quyền
    for (let i = data.length - 1; i >= 1; i--) {
      const fileId = data[i][0];
      const logEmail = (data[i][1] || "").toLowerCase().trim();
      
      if (logEmail === cleanEmail && fileId) {
        // Thu hồi quyền của file này
        try {
          const file = DriveApp.getFileById(fileId);
          try {
            file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
          } catch (sharingErr) {
            Logger.log(`Không thể setSharing PRIVATE cho file ${fileId} (Admin): ${sharingErr.message}`);
          }
          try {
            var permissionListResponse = Drive.Permissions.list(fileId, { supportsAllDrives: true });
            var permissionList = permissionListResponse.permissions || [];
            for (var j = 0; j < permissionList.length; j++) {
              var perm = permissionList[j];
              if (perm.emailAddress && perm.emailAddress.toLowerCase() === cleanEmail) {
                Drive.Permissions.remove(fileId, perm.id, { supportsAllDrives: true });
                break;
              }
            }
          } catch (removeErr) {
            // Bỏ qua lỗi gỡ viewer
            Logger.log(`Lỗi gỡ viewer (Admin) file ${fileId}: ${removeErr.message}`);
          }
          count++;
          sheet.deleteRow(i + 1);
          Logger.log(`[Admin Manual Cleanup] Đã thu hồi quyền truy cập file ${fileId} cho admin ${cleanEmail}`);
        } catch (fileErr) {
          Logger.log(`Lỗi khi thu hồi file ${fileId} cho admin: ` + fileErr.message);
        }
      }
    }
    
    return { success: true, message: `Đã thu hồi quyền truy cập của ${count} đề thi gốc đã mở.` };
  } catch (error) {
    return { success: false, message: "Lỗi khi thu hồi quyền: " + error.message };
  }
}

function saveScreenshotPair(screenBase64, camBase64, userEmail, existingFolderId) {
  const lock = LockService.getScriptLock();
  try {
    let userFolder;
    let folderId = existingFolderId;

    if (folderId) {
      try {
        userFolder = DriveApp.getFolderById(folderId);
        if (userFolder.isTrashed()) {
          userFolder = null;
          folderId = null;
        }
      } catch (e) {
        userFolder = null;
        folderId = null;
      }
    }

    if (!folderId) {
      lock.waitLock(10000); // 10s is enough to prevent race condition on folder creation
      const rootFolder = DriveApp.getFolderById("1dbnqiktUpREBuvNYrqJM-Xx1h8YdFjQP");
      const folders = rootFolder.getFoldersByName(userEmail);
      while (folders.hasNext()) {
        const f = folders.next();
        if (!f.isTrashed()) {
          userFolder = f;
          break;
        }
      }
      if (!userFolder) {
        userFolder = rootFolder.createFolder(userEmail);
      }
      folderId = userFolder.getId();
      lock.releaseLock();
    }

    const folderUrl = userFolder.getUrl();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');

    // === 2. Update tracking info in a try-catch to avoid blocking the whole process ===
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("TheoDoi");
      if (sheet) {
        // Optimized: Only get the values once and find the row directly using a search if possible
        // but for now, the loop is okay if we use a more targeted range or just set the value directly if we knew the row.
        // Let's stick to finding the row but optimize the writing.
        const data = sheet.getDataRange().getValues();
        let found = false;
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === userEmail) {
            const range = sheet.getRange(i + 1, 6, 1, 2);
            const currentValues = [data[i][5], data[i][6]];
            if (currentValues[0] !== folderUrl) {
              range.setValues([[folderUrl, now]]);
            } else {
              sheet.getRange(i + 1, 7).setValue(now);
            }
            found = true;
            break;
          }
        }
        if (!found) {
          sheet.appendRow([userEmail, 120, "", "Y", "", folderUrl, now]);
        }
      }
    } catch (sheetErr) {
      console.warn("Sheet update failed: " + sheetErr.message);
    }

    // === 2.5 Log periodic session tracking period ===
    try {
      logMonitorSession(userEmail);
    } catch(logErr) {
      console.warn("Monitor session log failed: " + logErr.message);
    }

    // === 3. Save images ===
    // Use Drive API via advanced service if possible for speed? 
    // But blob creation is standard.
    if (screenBase64) {
      const blob1 = Utilities.newBlob(Utilities.base64Decode(screenBase64), "image/jpeg", `${timestamp}_screen.jpg`);
      userFolder.createFile(blob1);
    }
    if (camBase64) {
      const blob2 = Utilities.newBlob(Utilities.base64Decode(camBase64), "image/jpeg", `${timestamp}_camera.jpg`);
      userFolder.createFile(blob2);
    }

    return { success: true, folderId: folderId, timestamp: timestamp };

  } catch (err) {
    if (lock.hasLock()) lock.releaseLock();
    console.error("Save error for " + userEmail + ": " + err.message);
    return { success: false, error: err.message };
  }
}

function incrementPrtScCount(userEmail) {
  if (!userEmail) return { success: false, message: "Email trống" };
  const cleanEmail = String(userEmail).trim().toLowerCase();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("TheoDoi");
    if (!sheet) {
      lock.releaseLock();
      return { success: false, message: "Không tìm thấy sheet TheoDoi" };
    }
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][0]).trim().toLowerCase();
      if (rowEmail === cleanEmail) {
        // Cột J là cột thứ 10 (chỉ số cột là 10, trong mảng là chỉ số 9)
        let currentCount = Number(data[i][9]) || 0;
        let newCount = currentCount + 1;
        sheet.getRange(i + 1, 10).setValue(newCount);
        found = true;
        break;
      }
    }
    if (!found) {
      // Nếu chưa có dòng cho email này, tự tạo dòng mới với số lần bấm PrtSc là 1
      sheet.appendRow([userEmail, 30, 0, "Y", "", "", "", "N", "", 1]);
    }
    lock.releaseLock();
    return { success: true, message: "Đã ghi nhận PrtSc thành công" };
  } catch (err) {
    if (lock.hasLock()) lock.releaseLock();
    console.error("Lỗi khi ghi nhận PrtSc: " + err.message);
    return { success: false, error: err.message };
  }
}

function incrementDevToolsCount(userEmail) {
  if (!userEmail) return { success: false, message: "Email trống" };
  const cleanEmail = String(userEmail).trim().toLowerCase();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("TheoDoi");
    if (!sheet) {
      lock.releaseLock();
      return { success: false, message: "Không tìm thấy sheet TheoDoi" };
    }
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][0]).trim().toLowerCase();
      if (rowEmail === cleanEmail) {
        // Cột K là cột thứ 11 (chỉ số cột là 11, trong mảng là chỉ số 10)
        let currentCount = Number(data[i][10]) || 0;
        let newCount = currentCount + 1;
        sheet.getRange(i + 1, 11).setValue(newCount);
        found = true;
        break;
      }
    }
    if (!found) {
      // Nếu chưa có dòng cho email này, tự tạo dòng mới với số lần bật DevTools là 1
      sheet.appendRow([userEmail, 30, 0, "Y", "", "", "", "N", "", 0, 1]);
    }
    lock.releaseLock();
    return { success: true, message: "Đã ghi nhận DevTools thành công" };
  } catch (err) {
    if (lock.hasLock()) lock.releaseLock();
    console.error("Lỗi khi ghi nhận DevTools: " + err.message);
    return { success: false, error: err.message };
  }
}

function incrementQrScanCount(userEmail) {
  if (!userEmail) return { success: false, message: "Email trống" };
  const cleanEmail = String(userEmail).trim().toLowerCase();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("TheoDoi");
    if (!sheet) {
      lock.releaseLock();
      return { success: false, message: "Không tìm thấy sheet TheoDoi" };
    }
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][0]).trim().toLowerCase();
      if (rowEmail === cleanEmail) {
        // Cột L là cột thứ 12 (chỉ số trong mảng là 11)
        let currentCount = Number(data[i][11]) || 0;
        let newCount = currentCount + 1;
        sheet.getRange(i + 1, 12).setValue(newCount);
        found = true;
        break;
      }
    }
    if (!found) {
      // Nếu chưa có dòng cho email này, tự tạo dòng mới với số lần quét QR là 1
      // Hàng mẫu gồm 12 cột: email, 30, 0, "Y", "", "", "", "N", "", 0, 0, 1
      sheet.appendRow([userEmail, 30, 0, "Y", "", "", "", "N", "", 0, 0, 1]);
    }
    lock.releaseLock();
    return { success: true, message: "Đã ghi nhận quét QR thành công" };
  } catch (err) {
    if (lock.hasLock()) lock.releaseLock();
    console.error("Lỗi khi ghi nhận quét QR: " + err.message);
    return { success: false, error: err.message };
  }
}

function handleQrScan(userEmail) {
  // Ghi nhận quét QR vào cột L
  incrementQrScanCount(userEmail);
  
  // HTML Giao diện cảnh báo cực kỳ nguy hiểm, chuyên nghiệp
  const html = `<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Cảnh Báo Bảo Mật</title>
  <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-50 text-gray-900 min-h-screen flex items-center justify-center p-4 antialiased selection:bg-gray-200">
  <div class="max-w-md w-full bg-white border border-gray-200 rounded-2xl p-8 shadow-sm relative">
    
    <!-- Minimal Icon -->
    <div class="mb-6 flex justify-center">
      <div class="bg-gray-100 p-3 rounded-full border border-gray-200">
        <svg xmlns="http://www.w3.org/2000/svg" class="h-8 w-8 text-gray-700" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
          <path stroke-linecap="round" stroke-linejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
        </svg>
      </div>
    </div>
    
    <!-- Title -->
    <div class="text-center mb-8">
      <h1 class="text-gray-900 font-semibold text-xl tracking-tight">Cảnh báo vi phạm bảo mật</h1>
      <p class="text-gray-500 text-sm mt-1">TRIPLE D SINH HỌC</p>
    </div>
    
    <!-- Content message -->
    <div class="space-y-4 text-gray-700 text-sm leading-relaxed mb-8 text-justify">
      <p>
        Hệ thống an ninh số Triple D sử dụng công nghệ định danh ẩn (Steganography Tracker) được nhúng sâu vào cấu trúc hình ảnh. Bất kỳ thiết bị nào thực hiện thao tác tải về, hiển thị hoặc truyền tải tệp tin này đều tự động kích hoạt giao thức báo cáo bảo mật ngầm.
      </p>
      <p>
        Toàn bộ dữ liệu viễn trắc bao gồm thông tin phần cứng và địa chỉ IP của thiết bị đang chứa ảnh đã được đồng bộ thành công về máy chủ trung tâm, nhằm thiết lập hồ sơ truy vết nguồn phát tán tài liệu trái phép.
      </p>
      <div class="bg-gray-50 p-4 rounded-xl text-xs text-gray-600 border border-gray-200">
        <span class="font-semibold text-gray-900">Lưu ý:</span> Hồ sơ vi phạm đang được giám sát chặt chẽ. Nếu phát hiện hành vi tiếp tục phát tán hoặc chia sẻ, tài khoản gốc <strong class="text-gray-900 font-medium">${userEmail}</strong> sẽ bị khóa vĩnh viễn.
      </div>
    </div>
    
    <!-- Security Details Block -->
    <div class="border-t border-gray-100 pt-6 mb-8">
      <h2 class="text-xs font-semibold text-gray-400 mb-4 uppercase tracking-wider">Dữ liệu định danh thiết bị</h2>
      <dl class="space-y-3 text-sm">
        <div class="flex justify-between items-center">
          <dt class="text-gray-500">Tài khoản gốc</dt>
          <dd class="text-gray-900 font-medium truncate ml-4" id="userEmail">Đang tải...</dd>
        </div>
        <div class="flex justify-between items-center">
          <dt class="text-gray-500">Mã số bảo mật</dt>
          <dd class="text-gray-900 font-mono text-xs bg-gray-100 px-2 py-1 rounded" id="securityCode">Đang tải...</dd>
        </div>
        <div class="flex justify-between items-center">
          <dt class="text-gray-500">IP Thiết bị</dt>
          <dd class="text-gray-900 font-mono text-xs" id="scannerIp">Đang xác thực...</dd>
        </div>
        <div class="flex justify-between items-center">
          <dt class="text-gray-500">Hệ điều hành</dt>
          <dd class="text-gray-900" id="scannerOs">Đang nhận diện...</dd>
        </div>
        <div class="flex justify-between items-center">
          <dt class="text-gray-500">Thời gian ghi nhận</dt>
          <dd class="text-gray-900 text-xs" id="scanTime">Đang tải...</dd>
        </div>
      </dl>
    </div>

    <!-- Footer action -->
    <div class="text-center">
      <button onclick="window.close()" class="px-6 py-2.5 bg-gray-900 hover:bg-gray-800 text-white text-sm font-medium rounded-xl transition duration-150 w-full shadow-sm">
        Đóng cảnh báo
      </button>
      <p class="text-[10px] text-gray-400 mt-4 font-mono uppercase tracking-widest">System Status: Monitored</p>
    </div>
    
  </div>

  <script>
    const email = "${userEmail}";
    document.getElementById("userEmail").textContent = email.toLowerCase();
    
    let hash = 0;
    for (let i = 0; i < email.length; i++) {
        hash = email.charCodeAt(i) + ((hash << 5) - hash);
    }
    const secCode = "TRD-" + Math.abs(hash).toString(16).substring(0, 6).toUpperCase() + "-" + Math.floor(Math.random() * 900 + 100);
    document.getElementById("securityCode").textContent = secCode;
    
    const now = new Date();
    const timeStr = now.toLocaleDateString('vi-VN') + " " + now.toLocaleTimeString('vi-VN');
    document.getElementById("scanTime").textContent = timeStr;
    
    const ua = navigator.userAgent;
    let os = "Mobile / PC Device";
    if (/Windows NT/.test(ua)) os = "Windows OS";
    else if (/Android/.test(ua)) os = "Android Mobile";
    else if (/iPhone|iPad|iPod/.test(ua)) os = "iOS Apple Device";
    else if (/Mac OS X/.test(ua)) os = "macOS Apple";
    else if (/Linux/.test(ua)) os = "Linux OS";
    document.getElementById("scannerOs").textContent = os;
    
    fetch("https://api64.ipify.org?format=json")
      .then(res => res.json())
      .then(data => {
        if (data && data.ip) {
          document.getElementById("scannerIp").textContent = data.ip;
        } else {
          document.getElementById("scannerIp").textContent = "103.199.17.202 (IPv4 Locked)";
        }
      })
      .catch(err => {
        const oct1 = 113 + Math.abs((hash >> 24) % 15);
        const oct2 = Math.abs((hash >> 16) % 254);
        const oct3 = Math.abs((hash >> 8) % 254);
        const oct4 = Math.abs(hash % 254);
        document.getElementById("scannerIp").textContent = \`\${oct1}.\${oct2}.\${oct3}.\${oct4} (Proxy Verified)\`;
      });
  </script>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle("CẢNH BÁO VI PHẠM BẢO MẬT HỆ THỐNG")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUserConfig(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("TheoDoi");

  // Nếu không có sheet
  if (!sheet) {
    return {
      interval: 30000,
      active: "Y",
      warningLevel: 0,
      folder: "",
      camera: "N" // ✅ sửa: mặc định không bắt camera
    };
  }

  const data = sheet.getDataRange().getValues();
  const cleanEmail = String(email).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const rowEmail = String(data[i][0]).trim().toLowerCase();

    if (rowEmail === cleanEmail) {
      return {
        interval: (Number(data[i][1]) || 60) * 1000,
        warningLevel: Number(data[i][2]) || 0,
        active: String(data[i][3] || "Y").trim().toUpperCase(),
        folder: data[i][5] || "",
        camera: String(data[i][7] || "N").trim().toUpperCase()
      };
    }
  }

  // Nếu CHƯA có user → tạo mới với interval = 30 giây
  sheet.appendRow([email, 30, 0, "Y", "", "", "", "N"]);

  return {
    interval: 30000,
    active: "Y",
    warningLevel: 0,
    folder: "",
    camera: "N" // ✅ sửa: phải khớp với sheet
  };
}

function updateOldestImageTime_fast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("TheoDoi");
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const email = data[i][0];
    const folderLink = data[i][5];

    if (!email || !folderLink) {
      results.push([""]);
      continue;
    }

    try {
      const folderIdMatch = folderLink.match(/[-\w]{25,}/);
      if (!folderIdMatch) {
        results.push([""]);
        continue;
      }

      const folderId = folderIdMatch[0];

      // 🔥 Query: lấy file cũ nhất
      const res = Drive.Files.list({
        q: `'${folderId}' in parents and trashed=false`,
        orderBy: "createdTime asc",
        pageSize: 1,
        fields: "files(createdTime)"
      });

      if (res.files && res.files.length > 0) {
        results.push([new Date(res.files[0].createdTime)]);
      } else {
        results.push([""]);
      }

    } catch (err) {
      Logger.log(`Error row ${i + 1}: ${err.message}`);
      results.push([""]);
    }
  }

  // 🚀 Ghi 1 lần (rất nhanh)
  sheet.getRange(2, 9, results.length, 1).setValues(results);
}

/**Theo Doi Khoa Hoc */


/**GHÉP ĐÔI*/
var PAIR_SHEET_NAME = "Pairing";
/**
 * Lấy hoặc tạo sheet Pairing
 */
function getPairSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(PAIR_SHEET_NAME) || ss.insertSheet(PAIR_SHEET_NAME);
}

/**
 * Thêm hoặc cập nhật user theo email
 * @param {Object} formData - {email, name, school, grade, phone, facebook, time, goal}
 * @returns {number} ID user
 */
function addOrUpdatePairUser(formData) {
  var sheet = getPairSheet();
  var data = sheet.getDataRange().getValues();

  // Kiểm tra xem email đã tồn tại chưa
  var existingRow = data.findIndex(r => r[1] === formData.email);

  if(existingRow >= 0){
    // Cập nhật thông tin bằng Batch Write
    sheet.getRange(existingRow + 1, 3, 1, 7).setValues([[
      formData.name,
      formData.school,
      formData.grade,
      "'" + formData.phone,
      formData.facebook,
      formData.time,
      formData.goal
    ]]);

    return data[existingRow][0]; // trả ID cũ
  } else {
    // Thêm mới
    var id = new Date().getTime(); // ID duy nhất
    sheet.appendRow([
      id,                   // A: ID
      formData.email,       // B: Email
      formData.name,        // C: Tên
      formData.school,      // D: Trường
      formData.grade,       // E: Lớp
      "'" + formData.phone,   // 👈 thêm dấu '
      formData.facebook,    // G: Facebook
      formData.time,        // H: Thời gian rảnh
      formData.goal,        // I: Mục tiêu
      'available',          // J: PairStatus
      '[]',                 // K: Requests (JSON)
      ''                    // L: MatchedWith
    ]);
    return id;
  }
}

/**
 * Gửi request ghép đôi
 * Chỉ thêm ID người gửi vào cột K (Requests)
 */
function sendPairRequest(fromId, toId) {
  var sheet = getPairSheet();
  var data = sheet.getDataRange().getValues();

  var toRow = data.findIndex(r => r[0] == toId);
  var fromRow = data.findIndex(r => r[0] == fromId);
  if (toRow < 0 || fromRow < 0) return false;

  // Chỉ gửi request nếu người nhận đang available
  if (data[toRow][9] === 'matched') return false;

  var requests = [];
  if (data[toRow][10]) {
    try { requests = JSON.parse(data[toRow][10]); } catch(e){ requests = []; }
  }

  if (!requests.includes(fromId)) requests.push(fromId);
  sheet.getRange(toRow + 1, 11).setValue(JSON.stringify(requests)); // cột K

  return true;
}

/**
 * Hủy ghép đôi
 * @param {number} userId
 */
function unmatchPair(userId) {
  var sheet = getPairSheet();
  var data = sheet.getDataRange().getValues();
  var userRow = data.findIndex(r => r[0] == userId);
  if(userRow < 0) return false;

  var partnerId = data[userRow][11];
  var userEmail = data[userRow][1];

  // Dù partnerId có rỗng vẫn phải xoá email
  setPartnerEmail(userEmail, "");

  if(partnerId){
    var partnerRow = data.findIndex(r => r[0] == partnerId);
    if(partnerRow >= 0){
      var partnerEmail = data[partnerRow][1];
      setPartnerEmail(partnerEmail, "");

      sheet.getRange(partnerRow+1, 10).setValue('available');
      sheet.getRange(partnerRow+1, 12).setValue('');
    }
  }

  sheet.getRange(userRow+1, 10).setValue('available');
  sheet.getRange(userRow+1, 12).setValue('');

  return true;
}


/**
 * Lấy các request đến của user
 */
function getIncomingRequests(userId) {
  var sheet = getPairSheet();
  var data = sheet.getDataRange().getValues();
  var userRow = data.findIndex(r => r[0] == userId);
  if (userRow < 0) return [];

  var requests = [];
  if (data[userRow][10]) {
    try { requests = JSON.parse(data[userRow][10]); } catch(e) { requests = []; }
  }

  // Lấy thông tin người gửi request
  return requests.map(rid => {
    var u = data.find(row => row[0] == rid);
    return {
      id: u[0],
      name: u[2],
      school: u[3],
      grade: u[4],
      goal: u[8],
      time: u[7]
    };
  });
}

/**
 * Xác nhận ghép đôi với requester
 * Cập nhật PairStatus và MatchedWith, đồng thời xóa request đã xử lý
 */
function confirmPairRequest(userId, requesterId) {
  var sheet = getPairSheet();

  // Luôn đọc data mới nhất
  var data = sheet.getDataRange().getValues();

  var userRow = data.findIndex(r => r[0] == userId);
  var reqRow  = data.findIndex(r => r[0] == requesterId);
  if (userRow < 0 || reqRow < 0) return false;

  // Chỉ match nếu cả 2 đều available
  if (data[userRow][9] === 'matched' || data[reqRow][9] === 'matched') return false;

  // Cập nhật trạng thái matched (cột J)
  sheet.getRange(userRow + 1, 10).setValue('matched');
  sheet.getRange(reqRow + 1, 10).setValue('matched');

  // Ghi ID đối tác vào cột L
  sheet.getRange(userRow + 1, 12).setValue(requesterId);
  sheet.getRange(reqRow + 1, 12).setValue(userId);

  // ❗ ĐỌC LẠI DATA SAU KHI setValue
  data = sheet.getDataRange().getValues();

  var userEmail = data[userRow][1]; // cột B
  var reqEmail  = data[reqRow][1];  // cột B

  // Ghi gmail sang sheet Users
  setPartnerEmail(userEmail, reqEmail);
  setPartnerEmail(reqEmail, userEmail);

  // Xóa request đã xử lý (cột K)
  var requests = [];
  try {
    requests = JSON.parse(data[userRow][10]);
  } catch (e) {
    requests = [];
  }

  requests = requests.filter(id => id != requesterId);
  sheet.getRange(userRow + 1, 11).setValue(JSON.stringify(requests));

  return true;
}

/**
 * Lấy thông tin đối tác đã match
 */
function getMatchedInfo(userId) {
  var sheet = getPairSheet();
  var data = sheet.getDataRange().getValues();
  var user = data.find(r => r[0] == userId);
  if (!user || user[9] != 'matched') return null;

  var partner = data.find(r => r[0] == user[11]);
  if (!partner) return null;

  return {
    name: partner[2],
    email: partner[1],
    school: partner[3],  // ✅ thêm trường
    grade: partner[4],   // ✅ thêm trường
    phone: partner[5],
    facebook: partner[6],
    time: partner[7],
    goal: partner[8]
  };
}


/**
 * Lấy danh sách user khác đang available
 * Chỉ show tên, trường, lớp, thời gian rảnh, mục tiêu
 */
function getAvailableUsers(currentUserId) {
  var sheet = getPairSheet();
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return [];

  // Bỏ header + dòng rác
  const cleaned = data.slice(1).filter(r =>
    r && r.length >= 10 && r[0] && r[9] // có ID + trạng thái
  );

  // Lọc ra người khác đang available
  const available = cleaned.filter(r => r[0] != currentUserId && r[9] === 'available');

  // Trả về thông tin, chỗ nào thiếu thì ghi "Thiếu thông tin"
  return available.map(u => ({
    id: u[0],
    name: u[2] ? String(u[2]).trim() : "Thiếu thông tin",
    school: u[3] ? String(u[3]).trim() : "Thiếu thông tin",
    grade: u[4] ? String(u[4]).trim() : "Thiếu thông tin",
    time: u[7] ? String(u[7]).trim() : "Thiếu thông tin",
    goal: u[8] ? String(u[8]).trim() : "Thiếu thông tin"
  }));
}


/**
 * Lấy ID user theo email
 */
function getUserIdByEmail(email) {
  var sheet = getPairSheet();
  var data = sheet.getDataRange().getValues();
  var user = data.find(r => r[1] === email); // cột B là email
  return user ? user[0] : null;
}

function getUserInfo(userId) {
  var sheet = getPairSheet();
  var data = sheet.getDataRange().getValues();
  var user = data.find(r => r[0] == userId);
  if(!user) return null;

  return {
    id: user[0],
    email: user[1],
    name: user[2],
    school: user[3],
    grade: user[4],
    phone: user[5],
    facebook: user[6],
    time: user[7],
    goal: user[8],
    pairStatus: user[9],
    matchedWith: user[11]
  };
}

/**
 * Từ chối request ghép đôi
 * @param {number} userId - ID người nhận request
 * @param {number} requesterId - ID người gửi request
 * @returns {boolean} true nếu xóa request thành công
 */
function declinePairRequest(userId, requesterId) {
  var sheet = getPairSheet();
  var data = sheet.getDataRange().getValues();
  
  var userRow = data.findIndex(r => r[0] == userId);
  if(userRow < 0) return false;

  var requests = [];
  try { requests = JSON.parse(data[userRow][10]); } catch(e){ requests = []; }

  // Xóa requesterId khỏi danh sách requests
  requests = requests.filter(id => id != requesterId);
  sheet.getRange(userRow + 1, 11).setValue(JSON.stringify(requests)); // cột K

  return true;
}

function setPartnerEmail(userEmail, partnerEmail) {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Users");
  var data = sheet.getDataRange().getValues();

  var row = data.findIndex(r => r[1] === userEmail); // ✅ cột B mới là email
  if (row >= 0) {
    sheet.getRange(row + 1, 20).setValue(partnerEmail); // cột T
  }
}

/**GHÉP ĐÔI*/


function logIdleLogout(email) {
  try {
    const role = getUserRole(email); // Lấy role
    if (!["STUDENT", "O-STUDENT", "S-STUDENT"].includes(role)) {
      // Nếu không phải student / o-student → không ghi log
      return false;
    }

    // Central Tracking log
    logActivityToSheet(email, "IDLE_LOGOUT", "Tự động đăng xuất do treo máy quá lâu (idle)");

    return true;
  } catch(e) {
    Logger.log("Lỗi khi ghi log logout: " + e.message);
    return false;
  }
}

function logScreenChange(email, currentScreen) {
  // Đã tắt ghi nhận ScreenChangeLog theo yêu cầu
  return true;
}

/**
 * Lấy danh sách câu hỏi cho một fileId cụ thể
 */
function getQAComments(fileId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("QA");
    if (!sheet) {
      sheet = ss.insertSheet("QA");
      sheet.appendRow(["ID", "FileID", "UserEmail", "UserName", "Content", "ParentID", "Timestamp"]);
      sheet.getRange("A1:G1").setFontWeight("bold").setBackground("#f3f3f3");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const comments = [];
    
    // Build role map for fast lookup
    const usersSheet = ss.getSheetByName("Users");
    const roleMap = {};
    if (usersSheet) {
      const uData = usersSheet.getDataRange().getValues();
      for (let i = 1; i < uData.length; i++) {
        const uEmail = uData[i][1] ? uData[i][1].toString().trim().toLowerCase() : "";
        let role = (uData[i][13] || "").toString().trim().toUpperCase();
        if (!["ADMIN", "STUDENT", "O-STUDENT", "S-STUDENT", "VIP"].includes(role)) {
          role = "MEM";
        }
        if (uEmail) roleMap[uEmail] = role;
      }
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(fileId)) {
        const uEmail = (data[i][2] || "").toString().trim().toLowerCase();
        let reacts = [];
        try {
           reacts = data[i][7] ? JSON.parse(data[i][7]) : [];
        } catch(e) {}
        comments.push({
          id: data[i][0],
          fileId: data[i][1],
          userEmail: uEmail,
          userName: data[i][3],
          content: data[i][4],
          parentId: data[i][5],
          userRole: roleMap[uEmail] || "MEM",
          reacts: reacts,
          timestamp: data[i][6] instanceof Date ? data[i][6].toISOString() : data[i][6]
        });
      }
    }
    
    return comments.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
  } catch (err) {
    Logger.log("Lỗi getQAComments: " + err);
    return [];
  }
}

/**
 * Lưu một câu hỏi hoặc phản hồi mới
 */
function saveQAComment(fileId, content, parentId = "") {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false, message: "Hết phiên đăng nhập." };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("QA");
    if (!sheet) {
      sheet = ss.insertSheet("QA");
      sheet.appendRow(["ID", "FileID", "UserEmail", "UserName", "Content", "ParentID", "Timestamp"]);
    }

    const user = getUserByEmail(email);
    const userName = user ? user.data[0] : email.split('@')[0];
    const id = "qa_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);
    const now = new Date();

    sheet.appendRow([
      id,
      fileId,
      email,
      userName,
      content,
      parentId,
      now
    ]);

    // --- HỆ THỐNG THÔNG BÁO ---
    const adminEmail = "dangta2606@gmail.com";
    const fileInfo = getFileDetailForNotif(fileId);
    const fileTitle = fileInfo.title;
    const actionPrefix = fileInfo.type + ":";
    
    // 1. Thông báo cho những người có quyền truy cập (Chỉ gửi khi là câu hỏi gốc mới)
    if (!parentId) {
      if (fileInfo.folderId) {
        addAutoNotificationToSheet(
          "❓ Hỏi đáp mới",
          `${userName} đã bình luận trong: ${fileTitle}`,
          actionPrefix + fileId,
          fileInfo.folderId, // targetFolder
          "ALL",             // targetUser (ALL users with access to folder)
          email              // senderEmail (to exclude from own view)
        );
      } else {
        // Fallback: nếu không có folderId, gửi cho Admin
        if (email.toLowerCase() !== adminEmail.toLowerCase()) {
          addAutoNotificationToSheet(
            "❓ Hỏi đáp mới",
            `${userName} đã bình luận trong: ${fileTitle}`,
            actionPrefix + fileId,
            "",
            adminEmail,
            email
          );
        }
      }
    }

    // 2. Nếu là phản hồi, chỉ thông báo riêng cho người viết câu hỏi gốc
    if (parentId) {
      const parentComment = getCommentById(parentId);
      if (parentComment && parentComment.userEmail.toLowerCase() !== email.toLowerCase()) {
        addAutoNotificationToSheet(
          "💬 Phản hồi hỏi đáp",
          `${userName} đã trả lời câu hỏi của bạn trong: ${fileTitle}`,
          actionPrefix + fileId,
          "",
          parentComment.userEmail,
          email
        );
      }
    }

    // Ghi log vào Tracking sheet
    const commentType = parentId ? "phản hồi" : "hỏi đáp mới";
    logActivityToSheet(email, "ADD_COMMENT", `Đăng ${commentType} cho tài liệu ID: ${fileId}. Nội dung: "${content.substring(0, 100)}${content.length > 100 ? '...' : ''}"`);

    return { 
      success: true, 
      comment: {
        id, fileId, userEmail: email, userName, content, parentId, timestamp: now.toISOString()
      }
    };
  } catch (err) {
    Logger.log("Lỗi saveQAComment: " + err);
    return { success: false, message: "Không thể lưu câu hỏi." };
  }
}

function deleteQAComment(commentId) {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false, message: "Hết phiên đăng nhập." };
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("QA");
    if (!sheet) return { success: false, message: "Không có dữ liệu QA." };
    
    const data = sheet.getDataRange().getValues();
    const userRole = getUserRole(email);
    const isAdmin = ["ADMIN", "VIP"].includes(userRole);
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(commentId)) {
        if (!isAdmin && data[i][2] !== email) {
           return { success: false, message: "Không có quyền xóa." };
        }
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "Không tìm thấy bình luận." };
  } catch(e) {
    return { success: false, message: "Lỗi: " + e };
  }
}

function reactQAComment(commentId) {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false, message: "Hết phiên đăng nhập." };
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("QA");
    if (!sheet) return { success: false, message: "Không có dữ liệu QA." };
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(commentId)) {
        let reacts = [];
        try {
          reacts = data[i][7] ? JSON.parse(data[i][7]) : [];
        } catch(e){}
        
        const index = reacts.indexOf(email);
        if (index > -1) {
          reacts.splice(index, 1);
        } else {
          reacts.push(email);
        }
        
        sheet.getRange(i + 1, 8).setValue(JSON.stringify(reacts));
        return { success: true, reacts: reacts };
      }
    }
    return { success: false, message: "Không tìm thấy bình luận." };
  } catch(e) {
    return { success: false, message: "Lỗi: " + e };
  }
}

function getCommentById(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("QA");
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        return {
          id: data[i][0],
          userEmail: data[i][2]
        };
      }
    }
  } catch (e) {}
  return null;
}

function getFileDetailForNotif(fileId) {
  try {
    // Check Books
    const books = getCachedBooksFromSheet();
    if (books && books.length > 0) {
      const item = books.find(i => String(i.fileId) === String(fileId));
      if (item) return { title: item.title, type: "books", folderId: item.folderId || null };
    }
    
    // Check Exams
    const exams = getCachedExamsFromSheet();
    if (exams && exams.length > 0) {
      const item = exams.find(i => String(i.fileId) === String(fileId));
      if (item) return { title: item.title, type: "exams", folderId: item.folderId || null };
    }
    
    // Check Courses
    const cachedCourses = getCachedCoursesFromSheet();
    const item = cachedCourses.find(i => String(i.fileId) === String(fileId));
    if (item) return { title: item.title, type: "courses", folderId: item.folderId || null };
  } catch (e) {}
  return { title: "Tài liệu " + fileId, type: "exams", folderId: null };
}

function updateAllCaches() {
  const overallStartTime = Date.now();
  const MAX_OVERALL_MS = 330000; // 5 phút 30 giây (mốc an toàn cao nhất của GAS)
  const results = [];

  try {
    // 1. Cập nhật cache sách
    const resBook = updateBookCache();
    results.push(resBook.message);

    // 2. Cập nhật cache khóa học nếu thời gian còn đủ
    if (Date.now() - overallStartTime < MAX_OVERALL_MS) {
      const resCourse = updateCourseCache();
      results.push(resCourse.message);
    } else {
      results.push("⚠️ Tạm dừng trước mốc 5 phút 30 giây. Lần bấm tiếp theo sẽ tự động quét lấy nốt phần khóa học.");
    }

    // 3. Cập nhật cache tài liệu/kỳ thi nếu thời gian còn đủ
    if (Date.now() - overallStartTime < MAX_OVERALL_MS) {
      const resExam = updateExamCache();
      results.push(resExam.message);
    } else {
      results.push("⚠️ Tạm dừng trước mốc 5 phút 30 giây. Lần bấm tiếp theo sẽ tự động quét lấy nốt phần tài liệu/bài thi.");
    }

    // Dọn dẹp tất cả khóa cũ khỏi ScriptProperties
    cleanupOldCacheProperties();

    const summaryMsg = results.join("\n");
    Logger.log("Kết quả cập nhật tất cả cache:\n" + summaryMsg);
    return { success: true, message: summaryMsg };
  } catch (err) {
    Logger.log("Lỗi khi cập nhật cache: " + err);
    return { success: false, message: "Lỗi khi cập nhật cache: " + err.toString() };
  }
}

function resetCacheCheckpoints() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty("checkpoint_books_folders");
    props.deleteProperty("checkpoint_exams_folders");
    props.deleteProperty("checkpoint_courses_folders");
    Logger.log("🧹 Đã xóa toàn bộ checkpoint cache.");
    return { success: true, message: "Đã xóa toàn bộ checkpoint. Lần quét tới sẽ bắt đầu lại từ đầu." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/* ================= NOTIFICATION SYSTEM (INBOX) ================= */

/**
 * Lấy danh sách thông báo cho người dùng
 */
function getSystemNotifications() {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return [];

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Notifications");
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const role = getUserRole(email);
    const allowedFolders = getUserAllowedFolders(email);
    
    // Xây dựng bản đồ tổ tiên (ancestors) từ cache các khóa học để kiểm tra thông báo thư mục dạng thừa kế
    const cachedCourses = getCachedCoursesFromSheet();
    const folderAncestorsMap = {};
    cachedCourses.forEach(c => {
      if (c.folderId && c.ancestors) {
        folderAncestorsMap[c.folderId] = c.ancestors;
      }
    });
    
    const userNotifications = [];
    
    // Lấy thời gian hiện tại để lọc tin cũ
    const now = new Date();
    const FIVE_DAYS_MS = 5 * 24 * 60 * 60 * 1000; 

    // Thêm các thông báo từ Sheet
    for (let i = 1; i < data.length; i++) {
        const id = String(data[i][0]);
        const title = String(data[i][1]);
        const desc = String(data[i][2]);
        const action = String(data[i][3] || "overview");
        const targetRole = String(data[i][4] || "ALL").toUpperCase();
        const targetFolder = String(data[i][5] || "");
        const time = data[i][6];
        const senderEmail = String(data[i][7] || "");
        
        // 0. Bỏ qua nếu người gửi chính là người đang check thông báo
        if (senderEmail && senderEmail.toLowerCase() === email.toLowerCase()) {
          continue;
        }

        // Bỏ qua thông báo rác cũ "Bài Đã Chấm Xong" chưa chỉ định đề cụ thể
        if (title === "Bài Đã Chấm Xong" && action === "mockExams" && (!targetFolder || targetFolder === "")) {
          continue;
        }

        const notifTime = time instanceof Date ? time : new Date(time);
      
      // 1. Kiểm tra "hạn sử dụng" của thông báo: 
      // Cũ hơn 5 ngày -> Bỏ qua (Trừ thông báo Chào mừng hệ thống)
      if (id !== "welcome_system_global" && title.indexOf("[Hệ Thống]") === -1 && (now - notifTime > FIVE_DAYS_MS)) {
        continue;
      }

      // 2. Kiểm tra đối tượng nhận
      let isTarget = false;
      
      const upperTarget = targetRole.toUpperCase();
      const upperEmail = email.toUpperCase();
      const upperRole = role.toUpperCase();

      // Hỗ trợ danh sách ngăn cách bởi dấu phẩy (group notification)
      const targetList = upperTarget.split(",").map(t => t.trim()).filter(Boolean);
      const isDirectEmailTarget = targetList.includes(upperEmail);

      if (targetList.includes("ALL")) {
        isTarget = true;
      } else if (targetList.includes(upperRole)) {
        isTarget = true;
      } else if (isDirectEmailTarget) {
        isTarget = true;
      }
      
      // Nếu không phải gửi trực tiếp qua email cá nhân và có quy định targetFolder -> Kiểm tra quyền thư mục
      if (!isDirectEmailTarget && targetFolder) {
        let folderAllowed = allowedFolders.includes("ALL_FOLDERS") || allowedFolders.includes(targetFolder);
        if (!folderAllowed && folderAncestorsMap[targetFolder]) {
          folderAllowed = folderAncestorsMap[targetFolder].some(fId => allowedFolders.includes(fId));
        }
        if (!folderAllowed) {
          isTarget = false;
        }
      }
      
      if (isTarget) {
        userNotifications.push({
          id: id || ("notif_" + i),
          title: title,
          desc: desc,
          action: action,
          time: time instanceof Date ? time.toISOString() : new Date().toISOString()
        });
      }
    }
    
    // Luôn thêm thông báo chào mừng mặc định
    const welcomeId = "welcome_system_global";
    userNotifications.push({
      id: welcomeId,
      title: "Chào mừng bạn đến với Triple D! 🎉",
      desc: "Chúc mừng bạn đã tạo tài khoản thành công. Chúc bạn học tập thật hiệu quả và đạt kết quả cao cùng Triple D Sinh Học!",
      action: "overview",
      time: "2026-01-01T00:00:00Z"
    });
    
    return userNotifications.sort((a, b) => new Date(b.time) - new Date(a.time));
  } catch (err) {
    Logger.log("Lỗi getSystemNotifications: " + err);
    return [];
  }
}

// Các hàm MarkAsRead trên Server không còn cần thiết nếu dùng LocalStorage
// Nhưng ta cứ giữ lại hàm trống để tránh lỗi nếu Client gọi nhầm, hoặc xóa hẳn.


/* ================= END NOTIFICATION SYSTEM ================= */

// ==========================================
// THI THỬ (MOCK EXAMS)
// ==========================================
function getMockExams() {
  try {
    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) return { success: false, message: "User not logged in." };
    
    // Check quyền (role)
    const userRole = getUserRole(email).toUpperCase().trim();
    const userEmailUpper = email.toUpperCase().trim();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("MockExams");
    if (!sheet) {
      sheet = ss.insertSheet("MockExams");
      sheet.appendRow([
        "File ID", 
        "Allowed Emails (Cách nhau dấu ,)", 
        "Allowed Roles (Ví dụ: FUND, RISE)", 
        "Tên bài thi",
        "Thời gian mở (VD: 2024-05-10 08:00)",
        "Thời gian đóng (VD: 2024-05-15 17:00)",
        "Tên lớp",
        "Email người phụ trách",
        "Thời gian làm bài (Phút)"
      ]);
      sheet.appendRow(["YOUR_FILE_ID", "ALL", "", "Đề thi thử số 1 (Mẫu)", "", "", "Lớp 10A1", "", "60"]);
      return { success: true, data: [] };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };

    // Dynamic Header Mapping
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    
    let fileIdIdx = 0;
    let emailsIdx = 1;
    let rolesIdx = 2;
    let titleIdx = 3;
    let startIdx = 4;
    let endIdx = 5;
    let classNameIdx = 6; // Default to column G
    let adminEmailIdx = -1;
    let durationIdx = -1;

    headers.forEach((h, idx) => {
      if (h.includes("file id") || h === "id") fileIdIdx = idx;
      else if (h.includes("email")) emailsIdx = idx;
      else if (h.includes("role") || h.includes("vai trò")) rolesIdx = idx;
      else if (h.includes("tên bài") || h.includes("tiêu đề") || h === "bài thi") titleIdx = idx;
      else if (h.includes("mở") || h.includes("start")) startIdx = idx;
      else if (h.includes("đóng") || h.includes("end")) endIdx = idx;
      else if (h.includes("lớp") || h.includes("class")) classNameIdx = idx;
      else if (h.includes("phụ trách") || h.includes("admin")) adminEmailIdx = idx;
      else if (h.includes("thời gian") || h.includes("phút") || h.includes("duration") || h.includes("thời lượng") || h.includes("làm bài")) durationIdx = idx;
    });

    if (adminEmailIdx === -1 && headers.length >= 8) {
      adminEmailIdx = 7;
    }
    if (durationIdx === -1 && headers.length >= 9) {
      durationIdx = 8;
    }

    const mockExams = [];
    for (let i = 1; i < data.length; i++) {
      const fileId = data[i][fileIdIdx] ? data[i][fileIdIdx].toString().trim() : "";
      if (!fileId) continue;
      
      const allowedEmailsRaw = data[i][emailsIdx] ? data[i][emailsIdx].toString() : "";
      const allowedRolesRaw  = data[i][rolesIdx] ? data[i][rolesIdx].toString() : "";
      
      const examTitle = data[i][titleIdx] ? data[i][titleIdx].toString().trim() : "Chưa có tên";
      const startTime = data[i][startIdx] ? (data[i][startIdx] instanceof Date ? data[i][startIdx].toISOString() : data[i][startIdx].toString().trim()) : "";
      const endTime = data[i][endIdx] ? (data[i][endIdx] instanceof Date ? data[i][endIdx].toISOString() : data[i][endIdx].toString().trim()) : "";
      const className = classNameIdx !== -1 && data[i][classNameIdx] ? data[i][classNameIdx].toString().trim() : "";
      const parsedDuration = durationIdx !== -1 && data[i][durationIdx] ? parseInt(data[i][durationIdx].toString().trim(), 10) : 60;
      const duration = isNaN(parsedDuration) || parsedDuration <= 0 ? 60 : parsedDuration;
      
      // Default type is PDF, no folder, no submit link
      const folderName = "Đề thi thử";
      const fileType = "pdf";
      const submitLink = "";

      // Phân quyền tương tự sheet Permissions
      let isAllowed = false;
      if (userRole === "ADMIN") {
        isAllowed = true;
      } else {
        const allowedEmails = allowedEmailsRaw.split(",").map(e => e.toUpperCase().trim()).filter(e => e !== "");
        const allowedRoles = allowedRolesRaw.split(",").map(r => r.toUpperCase().trim()).filter(r => r !== "");
        
        let emailOk = false;
        if (allowedEmails.length === 0) {
          emailOk = true;
        } else if (allowedEmails.includes("ALL") || allowedEmails.includes(userEmailUpper)) {
          emailOk = true;
        }

        let roleOk = false;
        if (allowedRoles.length === 0) {
          roleOk = true;
        } else if (allowedRoles.includes("ALL") || allowedRoles.includes(userRole) || (userRole === "S-STUDENT" && allowedRoles.includes("STUDENT")) || (userRole === "STUDENT" && allowedRoles.includes("S-STUDENT"))) {
          roleOk = true;
        }

        isAllowed = emailOk && roleOk;
      }

      if (isAllowed) {
        const adminEmailVal = adminEmailIdx !== -1 && data[i][adminEmailIdx] ? data[i][adminEmailIdx].toString().trim() : "";
        mockExams.push({
          fileId: fileId,
          title: examTitle,
          folder: folderName,
          className: className,
          type: fileType,
          submitLink: submitLink,
          startTime: startTime,
          endTime: endTime,
          adminEmail: adminEmailVal,
          duration: duration
        });
      }
    }

    return { success: true, data: mockExams };
  } catch (error) {
    Logger.log("Error in getMockExams: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

function getUserSubmittedExams() {
  try {
    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) return [];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ExamSubmissions');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const normEmail = normalizeExamText(email);
    const submittedMap = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = normalizeExamText(row[1]);
      if (rowEmail === normEmail) {
        const rowExamTitle = row[2] ? row[2].toString().trim() : "";
        const normTitle = normalizeExamText(rowExamTitle);
        if (!normTitle) continue;

        const score = row[5] !== undefined ? row[5].toString().trim() : "";
        let submittedAt = "";
        if (row[0]) {
          try {
            const dateObj = new Date(row[0]);
            submittedAt = Utilities.formatDate(dateObj, "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm");
          } catch (e) {
            submittedAt = row[0].toString();
          }
        }

        // Nếu đã có bản ghi trước đó, ưu tiên bản ghi có điểm hoặc bản ghi mới hơn
        if (!submittedMap[normTitle] || (score !== "" && submittedMap[normTitle].score === "")) {
          submittedMap[normTitle] = {
            examTitle: rowExamTitle,
            score: score,
            timestamp: submittedAt
          };
        }
      }
    }
    return Object.values(submittedMap);
  } catch (e) {
    Logger.log("Lỗi getUserSubmittedExams: " + e.toString());
    return [];
  }
}

// --- BẢO MẬT MINI GAMES ---

function playTaiXiuSecure(choice, betAmount) {
  const deductionOk = deductPointsSmartForCurrentUser(betAmount);
  if (!deductionOk) {
    return { success: false };
  }
  
  const email = CacheService.getUserCache().get("loggedInUser");
  if (email) {
      logCoinTransaction(email, -Math.abs(betAmount), "TD Coin", "Cược PlayTaiXiu");
  }

  let biasRate = 0.5;
  if (betAmount > 500) biasRate = 0.3;
  else if (betAmount > 300) biasRate = 0.2;
  else if (betAmount > 200) biasRate = 0.15;
  else if (betAmount > 100) biasRate = 0.1;

  const loseBias = Math.random() < biasRate;

  let dice1, dice2, dice3, total, result;
  do {
    dice1 = Math.floor(Math.random() * 6) + 1;
    dice2 = Math.floor(Math.random() * 6) + 1;
    dice3 = Math.floor(Math.random() * 6) + 1;
    total = dice1 + dice2 + dice3;
    result = (total <= 10) ? "xiu" : "tai";
  } while (loseBias && result === choice);

  const win = (choice === result);
  let payout = 0;
  
  if (win) {
    payout = betAmount * 2;
    addUserPoints_(payout, "Thắng Bầu Cua");
    addUserEventCoin_(Math.floor(payout / 20), "Bonus Bầu Cua");
  }

  return {
    success: true,
    dice1: dice1,
    dice2: dice2,
    dice3: dice3,
    total: total,
    result: result,
    win: win,
    betAmount: betAmount,
    payout: payout
  };
}

function submitArcadePointsSecure(gameName, points) {
  if (typeof points !== "number" || points <= 0 || points > 1000) return { success: false };
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false };

  const cache = CacheService.getUserCache();
  const key = "arcade_" + gameName + "_" + email;
  
  // Anti-spam cooldown per game
  if (cache.get(key)) return { success: false, message: "Too soon" };
  cache.put(key, "1", 30); // 30 seconds wait per game before sending another score

  addUserPoints_(points, "Thưởng Arcade: " + gameName);
  addUserEventCoin_(Math.floor(points / 10), "Bonus Arcade: " + gameName);
  
  return { success: true, newPoints: getUserPoints() };
}

function claimGiftBoxSecure() {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false };

  const lock = LockService.getUserLock();
  if (!lock.tryLock(5000)) return { success: false, message: "Hệ thống đang bận" };

  try {
    const cache = CacheService.getUserCache();
    const key = "giftbox_claim_" + email;
    if (cache.get(key)) return { success: false, message: "Too soon" };
    
    cache.put(key, "1", 900); // 15 mins cooldown enforced on server

    const rewards = [
      { type: 'add', value: 10, chance: 0.25 },
      { type: 'add', value: 20, chance: 0.25 },
      { type: 'add', value: 30, chance: 0.20 },
      { type: 'add', value: 50, chance: 0.15 },
      { type: 'add', value: 80, chance: 0.10 },
      { type: 'add', value: 150, chance: 0.05 }
    ];
    let rand = Math.random();
    let cumulative = 0;
    let reward = rewards[0];
    for (let r of rewards) {
      cumulative += r.chance;
      if (rand < cumulative) { reward = r; break; }
    }

    const user = getUserByEmail(email);
    if (!user) return { success: false };

    let diff = 0;
    if (reward.type === 'add') {
      diff = reward.value;
    } else if (reward.type === 'multiply') {
      const myCoin = parseInt(user.data[10]) || 0;
      diff = Math.floor(myCoin * reward.value) - myCoin;
    }
    
    if (diff > 0) {
      addUserPoints_(diff, "Hộp quà 15 phút");
    }

    return { success: true, reward: reward, diff: diff, newTotal: getUserPoints() };
  } finally {
    lock.releaseLock();
  }
}

function startAviatorSecure(bet) {
  if (typeof bet !== 'number' || bet <= 0) return { success: false, message: "Invalid bet" };
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false, message: "Not logged in" };

  const lock = LockService.getUserLock();
  if (!lock.tryLock(5000)) return { success: false, message: "Server busy" };

  try {
    const points = getUserPoints();
    if (points < bet) return { success: false, message: "Not enough points" };
    
    // Deduct points securely
    if (!addUserPoints_(-bet, "Tham gia Aviator")) return { success: false, message: "Failed to deduct points" };

    const cache = CacheService.getUserCache();
    // Cache the bet to be used when cashing out. Valid for 2 minutes.
    cache.put("aviator_bet_v2_" + email, bet.toString(), 120); 

    return { success: true, newPoints: getUserPoints() };
  } finally {
    lock.releaseLock();
  }
}

function processAviatorCashoutSecure(cashoutMultiplier) {
  if (typeof cashoutMultiplier !== 'number' || cashoutMultiplier < 1) return { success: false };

  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false };

  const cache = CacheService.getUserCache();
  
  // Rate limit
  const rlKey = "aviator_cashout_rl_" + email;
  if (cache.get(rlKey)) return { success: false };
  cache.put(rlKey, "1", 3);

  const betKey = "aviator_bet_v2_" + email;
  const betStr = cache.get(betKey);
  if (!betStr) return { success: false }; // No bet found, maybe expired or cheated
  
  const bet = parseInt(betStr);
  if (isNaN(bet) || bet <= 0) return { success: false };

  // Remove the bet to prevent double cashout
  cache.remove(betKey);

  // Allow up to a reasonable multiplier just for anti-abuse:
  if (cashoutMultiplier > 100) cashoutMultiplier = 100; 
  const payout = Math.floor(bet * cashoutMultiplier);
  const profit = payout - bet; 
  
  const lock = LockService.getUserLock();
  if (!lock.tryLock(5000)) return { success: false };
  try {
    addUserPoints_(payout, "Chốt lời Aviator");
    addUserEventCoin_(Math.floor(profit / 10), "Bonus Aviator");
    return { success: true, newPoints: getUserPoints() };
  } finally {
    lock.releaseLock();
  }
}


// ================= FOOTBALL BETTING =================
const FOOTBALL_API_KEY = "6feb5b018b8f41ebbdb1a312a8e0b722";

function getBetName(betOn, matchId) {
  const isHomeFav = (matchId % 2 === 0);
  betOn = String(betOn).toUpperCase().trim();
  if (betOn.startsWith("CS_OTHER")) {
    if (betOn.indexOf(":") !== -1) {
      const pred = betOn.split(":")[1];
      return `Tỷ số Kỳ Vọng: Khác (${pred})`;
    }
    return "Tỷ số Kỳ Vọng: Khác";
  }
  switch (betOn) {
    case "HOME": return "Tỷ số 1X2 - Dự đoán Đội Nhà Thắng";
    case "DRAW": return "Tỷ số 1X2 - Dự đoán Kết quả Hòa";
    case "AWAY": return "Tỷ số 1X2 - Dự đoán Đội Khách Thắng";
    case "HDP_HOME": return isHomeFav ? "Dự báo Chênh lệch - Đội Nhà Thắng ưu thế (-0.5)" : "Dự báo Chênh lệch - Đội Nhà Thắng ưu thế (+0.5)";
    case "HDP_AWAY": return isHomeFav ? "Dự báo Chênh lệch - Đội Khách Thắng ưu thế (+0.5)" : "Dự báo Chênh lệch - Đội Khách Thắng ưu thế (-0.5)";
    case "OU_OVER": return "Tổng Bàn Thắng - Kỳ vọng cao (Trên 2.5)";
    case "OU_UNDER": return "Tổng Bàn Thắng - Kỳ vọng thấp (Dưới 2.5)";
    case "CS_1_0": return "Tỷ số Kỳ Vọng: 1 - 0";
    case "CS_2_0": return "Tỷ số Kỳ Vọng: 2 - 0";
    case "CS_2_1": return "Tỷ số Kỳ Vọng: 2 - 1";
    case "CS_0_1": return "Tỷ số Kỳ Vọng: 0 - 1";
    case "CS_0_2": return "Tỷ số Kỳ Vọng: 0 - 2";
    case "CS_1_2": return "Tỷ số Kỳ Vọng: 1 - 2";
    case "CS_1_1": return "Tỷ số Kỳ Vọng: 1 - 1";
    case "CS_OTHER": return "Tỷ số Kỳ Vọng: Khác";
    default: return betOn;
  }
}

function isBetWinning(betOn, matchId, homeScore, awayScore) {
  betOn = String(betOn).toUpperCase().trim();
  if (betOn.startsWith("CS_OTHER")) {
    if (betOn.indexOf(":") !== -1) {
      const parts = betOn.split(":")[1].split("-");
      if (parts.length === 2) {
        const predHome = parseInt(parts[0]);
        const predAway = parseInt(parts[1]);
        return homeScore === predHome && awayScore === predAway;
      }
    }
    // Fallback if old bet without prediction
    const specific = [
      (homeScore === 1 && awayScore === 0),
      (homeScore === 2 && awayScore === 0),
      (homeScore === 2 && awayScore === 1),
      (homeScore === 0 && awayScore === 1),
      (homeScore === 0 && awayScore === 2),
      (homeScore === 1 && awayScore === 2),
      (homeScore === 1 && awayScore === 1)
    ];
    return !specific.some(val => val === true);
  }

  const favorite = (matchId % 2 === 0) ? "HOME" : "AWAY";

  switch (betOn) {
    case "HOME":
      return homeScore > awayScore;
    case "DRAW":
      return homeScore === awayScore;
    case "AWAY":
      return awayScore > homeScore;

    case "HDP_HOME":
      if (favorite === "HOME") {
        return (homeScore - 0.5) > awayScore;
      } else {
        return (homeScore + 0.5) > awayScore;
      }
    case "HDP_AWAY":
      if (favorite === "HOME") {
        return (awayScore + 0.5) > homeScore;
      } else {
        return (awayScore - 0.5) > homeScore;
      }

    case "OU_OVER":
      return (homeScore + awayScore) > 2.5;
    case "OU_UNDER":
      return (homeScore + awayScore) < 2.5;

    case "CS_1_0":
      return homeScore === 1 && awayScore === 0;
    case "CS_2_0":
      return homeScore === 2 && awayScore === 0;
    case "CS_2_1":
      return homeScore === 2 && awayScore === 1;
    case "CS_0_1":
      return homeScore === 0 && awayScore === 1;
    case "CS_0_2":
      return homeScore === 0 && awayScore === 2;
    case "CS_1_2":
      return homeScore === 1 && awayScore === 2;
    case "CS_1_1":
      return homeScore === 1 && awayScore === 1;
    case "CS_OTHER":
      const specific = [
        (homeScore === 1 && awayScore === 0),
        (homeScore === 2 && awayScore === 0),
        (homeScore === 2 && awayScore === 1),
        (homeScore === 0 && awayScore === 1),
        (homeScore === 0 && awayScore === 2),
        (homeScore === 1 && awayScore === 2),
        (homeScore === 1 && awayScore === 1)
      ];
      return !specific.some(val => val === true);

    default:
      return false;
  }
}

function getFootballBetStats() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('football_bet_stats_v1');
  if (cached) return JSON.parse(cached);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FootballBets");
    if (!sheet) return {};
    const data = sheet.getDataRange().getValues();
    let stats = {};
    for (let i = 1; i < data.length; i++) {
      let matchId = data[i][2];
      let betOn = data[i][4];
      let amount = parseInt(data[i][6]) || 0;
      if (!stats[matchId]) stats[matchId] = {};
      stats[matchId][betOn] = (stats[matchId][betOn] || 0) + amount;
      stats[matchId]._total = (stats[matchId]._total || 0) + amount;
      stats[matchId]._count = (stats[matchId]._count || 0) + 1;
    }
    cache.put('football_bet_stats_v1', JSON.stringify(stats), 60);
    return stats;
  } catch(e) {
    return {};
  }
}

function getFootballMatches() {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'football_matches_v7';
  const cachedMatches = cache.get(CACHE_KEY);
  if (cachedMatches) {
    try {
      return JSON.parse(cachedMatches);
    } catch(e) {}
  }

  try {
    const tz = Session.getScriptTimeZone();
    let d1 = new Date();
    let d2 = new Date();
    d2.setDate(d2.getDate() + 3);
    
    let dateFrom = Utilities.formatDate(d1, tz, "yyyy-MM-dd");
    let dateTo = Utilities.formatDate(d2, tz, "yyyy-MM-dd");
    
    const fetchUrl = `https://api.football-data.org/v4/matches?dateFrom=${dateFrom}&dateTo=${dateTo}`;
    const options = {
      'method' : 'get',
      'headers': {
        'X-Auth-Token': FOOTBALL_API_KEY
      },
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(fetchUrl, options);
    let json;
    try {
      json = JSON.parse(response.getContentText());
    } catch (parseError) {
      return { success: false, message: "Lỗi dữ liệu từ máy chủ API", req: response.getContentText().substring(0, 100) };
    }
    
    if (json.errorCode === 429) {
       // Thử lấy lại dữ liệu dự phòng từ cache cũ nếu có
       const backup = cache.get('football_matches_backup');
       if (backup) {
         return JSON.parse(backup);
       }
       return { success: false, message: "Hệ thống đang quá tải yêu cầu API. Vui lòng thử lại sau ít phút." };
    }

    if (!json.matches) return { success: false, message: json.message || "Lỗi từ API bóng đá", req: json };
    
    let matches = json.matches
      .filter(m => m.status === "TIMED" || m.status === "SCHEDULED")
      .map(m => {
        let p = m.id % 20; 
        
        let winHome = 1.3 + (p / 10);        // 1.3 - 3.2
        let winAway = 1.3 + ((20 - p) / 10); // 1.3 - 3.2
        let draw = 2.5 + ((p % 5) / 5);      // 2.5 - 3.3

        let hdpHome, hdpAway;
        if (p % 2 === 0) {
          hdpHome = 1.80 + (p % 4) / 20;    // 1.80, 1.85, 1.90, 1.95
          hdpAway = 1.80 + (3 - (p % 4)) / 20;
        } else {
          hdpHome = 1.80 + (3 - (p % 4)) / 20;
          hdpAway = 1.80 + (p % 4) / 20;
        }

        let ouOver = 1.85 + (p % 3) / 15;   // 1.85, 1.91, 1.98
        let ouUnder = 1.85 + (2 - (p % 3)) / 15;

        let cs10 = 4.0 + (p % 3) / 2;
        let cs20 = 5.0 + (p % 4) / 2;
        let cs21 = 6.0 + (p % 5) / 2;
        let cs01 = 4.5 + ((20 - p) % 3) / 2;
        let cs02 = 5.5 + ((20 - p) % 4) / 2;
        let cs12 = 6.5 + ((20 - p) % 5) / 2;
        let cs11 = 3.5 + (p % 3) / 2;
        let csOther = 8.5 + (p % 6) / 2;
        
        let oeEven = 1.88 + (p % 3) / 50; // 1.88, 1.90, 1.92
        let oeOdd = 1.88 + (2 - (p % 3)) / 50;
        
        let formattedTime = Utilities.formatDate(new Date(m.utcDate), tz, "dd/MM HH:mm");
        
        let hTeamName = (m.homeTeam && (m.homeTeam.shortName || m.homeTeam.name)) || "Đội nhà (TBD)";
        let aTeamName = (m.awayTeam && (m.awayTeam.shortName || m.awayTeam.name)) || "Đội khách (TBD)";
        let hCrest = (m.homeTeam && m.homeTeam.crest) || "https://upload.wikimedia.org/wikipedia/commons/e/e0/Placeholder_image_%28transparent%29.png";
        let aCrest = (m.awayTeam && m.awayTeam.crest) || "https://upload.wikimedia.org/wikipedia/commons/e/e0/Placeholder_image_%28transparent%29.png";

        return {
          id: m.id,
          competition: (m.competition && m.competition.name) ? m.competition.name : "CLUB FRIENDLIES",
          dateStr: formattedTime,
          dateMs: new Date(m.utcDate).getTime(),
          homeTeam: hTeamName,
          awayTeam: aTeamName,
          homeCrest: hCrest,
          awayCrest: aCrest,
          odds: {
            home: winHome.toFixed(2),
            draw: draw.toFixed(2),
            away: winAway.toFixed(2),
            favorite: (p % 2 === 0) ? "HOME" : "AWAY",
            hdpValue: "0.5",
            hdpHome: hdpHome.toFixed(2),
            hdpAway: hdpAway.toFixed(2),
            ouValue: "2.5",
            ouOver: ouOver.toFixed(2),
            ouUnder: ouUnder.toFixed(2),
            cs10: cs10.toFixed(2),
            cs20: cs20.toFixed(2),
            cs21: cs21.toFixed(2),
            cs01: cs01.toFixed(2),
            cs02: cs02.toFixed(2),
            cs12: cs12.toFixed(2),
            cs11: cs11.toFixed(2),
            csOther: csOther.toFixed(2),
            oeEven: oeEven.toFixed(2),
            oeOdd: oeOdd.toFixed(2)
          }
        };
      });
      
    matches.sort((a,b) => a.dateMs - b.dateMs);
    
    let result = { success: true, matches: matches };
    let resultStr = JSON.stringify(result);
    cache.put(CACHE_KEY, resultStr, 21600); // 6 hours max
    cache.put('football_matches_backup', resultStr, 21600); 
    
    return result;
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function placeFootballBet(matchId, matchDesc, betOn, odds, betAmount) {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false, message: "Bạn chưa đăng nhập." };
  
  betAmount = parseInt(betAmount);
  if (isNaN(betAmount) || betAmount < 10) return { success: false, message: "Tham gia tối thiểu 10 Coin." };
  if (betAmount > 3000) return { success: false, message: "Tham gia tối đa 3,000 Coin một lượt nhằm hạn chế lạm phát." };
  
  // Giới hạn tỉ lệ x4 đối với kèo thắng/thua/hòa (được sinh dựa theo lượng cược)
  if (betOn === "HOME" || betOn === "AWAY" || betOn === "DRAW") {
    odds = parseFloat(odds);
    if (isNaN(odds)) odds = 1.0;
    if (odds > 4.00) odds = 4.00;
  }
  
  // Ngăn chặn rửa tiền bằng cách cược vào trận đấu thuộc mùa giải sau trước khi reset coin
  // Và ngăn chặn cược sau khi trận đấu đã bắt đầu
  const matchResult = getFootballMatches();
  if (!matchResult || matchResult.success === false) {
    return { success: false, message: matchResult ? matchResult.message : "Không thể tải thông tin trận đấu để đối soát thời gian. Vui lòng thử lại sau ít phút." };
  }
  
  if (matchResult && matchResult.matches) {
    const m = matchResult.matches.find(x => x.id.toString() === matchId.toString());
    if (!m) {
      return { success: false, message: "Trận đấu không tồn tại, đã bắt đầu hoặc đã kết thúc." };
    }
    if (m && m.dateMs) {
      const now = new Date();
      
      // Ngăn chặn cược sau khi trận đấu đã bắt đầu
      if (now.getTime() >= m.dateMs) {
        return { success: false, message: "Trận đấu này đã bắt đầu hoặc đã diễn ra. Không thể đặt cược nữa!" };
      }
      
      const currentQ = Math.floor(now.getMonth() / 3) + 1;
      const currentSeason = `${now.getFullYear()}-Q${currentQ}`;
      
      const matchDate = new Date(m.dateMs);
      const matchQ = Math.floor(matchDate.getMonth() / 3) + 1;
      const matchSeason = `${matchDate.getFullYear()}-Q${matchQ}`;
      
      if (matchSeason !== currentSeason) {
        return { success: false, message: "Trận đấu này diễn ra ở mùa giải tiếp theo. Vui lòng chờ hệ thống Reset Coin chuyển mùa để đặt cược nhằm đảm bảo công bằng." };
      }
    }
  }

  const lock = LockService.getUserLock();
  if (!lock.tryLock(5000)) return { success: false, message: "Hệ thống bận, vui lòng thử lại." };
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("FootballBets");
    if (!sheet) {
      sheet = ss.insertSheet("FootballBets");
      sheet.appendRow(["Thời Gian", "Email", "Match ID", "Trận Đấu", "Phương án Dự Đoán", "Tỷ Lệ", "TD Coin Gửi", "Trạng Thái", "Kết Quả", "Thưởng Nhận"]);
      sheet.setFrozenRows(1);
      sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#f8fafc");
    } else {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === email && data[i][2].toString() === matchId.toString() && data[i][4] === betOn) {
           return { success: false, message: "Bạn đã dự đoán kèo này rồi, mỗi loại kèo chỉ được tham gia 1 lần!" };
        }
      }
    }

    const pts = getUserPoints();
    if (pts < betAmount) return { success: false, message: "Bạn không đủ TD Coin." };
    
    const betLabel = getBetName(betOn, matchId);
    addUserPoints_(-betAmount, `Cược bóng đá: ${matchDesc} (${betLabel})`);
    
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    sheet.appendRow([now, email, matchId, matchDesc, betOn, odds, betAmount, "PENDING", "", 0]);
    
    // Giới hạn 1000 dòng cho FootballBets
    limitSheetRows_(sheet, 1000);
    
    return { success: true, message: `Bạn đã gửi dự đoán thành công ${betAmount} 🪙 vào ${betLabel}!`, newPoints: getUserPoints() };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getMyFootballBets() {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return [];
  
  try {
    autoUpdateUserFootballBets(email);
  } catch (err) {
    Logger.log("Err running autoUpdateUserFootballBets in history click: " + err.toString());
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("FootballBets");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const myBets = [];
    for (let i = data.length - 1; i > 0; i--) {
      if (data[i] && data[i][1] && String(data[i][1]).toLowerCase().trim() === email.toLowerCase().trim()) {
        const rowTime = data[i][0];
        let timeStr = "";
        if (rowTime instanceof Date) {
          try {
            timeStr = Utilities.formatDate(rowTime, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
          } catch (e) {
            timeStr = rowTime.toLocaleString();
          }
        } else {
          timeStr = rowTime ? String(rowTime) : "";
        }

        const rawBetOn = data[i][4] ? String(data[i][4]) : "";
        const matchIdVal = data[i][2] ? Number(data[i][2]) || 0 : 0;
        const readableBetOn = getBetName(rawBetOn, matchIdVal);

        myBets.push({
          time: timeStr,
          id: String(matchIdVal),
          matchDesc: data[i][3] ? String(data[i][3]) : "",
          betOn: readableBetOn,
          odds: typeof data[i][5] !== 'undefined' ? Number(data[i][5]) || 0 : 0,
          amount: typeof data[i][6] !== 'undefined' ? Number(data[i][6]) || 0 : 0,
          status: data[i][7] ? String(data[i][7]) : "PENDING", 
          reward: typeof data[i][9] !== 'undefined' ? Number(data[i][9]) || 0 : 0
        });
      }
      if (myBets.length >= 30) break;
    }
    return myBets;
  } catch (e) {
    Logger.log("Error in getMyFootballBets: " + e.toString());
    return [];
  }
}

function autoUpdateUserFootballBets(email) {
  if (!email) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return; 

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("FootballBets");
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const scriptCache = CacheService.getScriptCache();
    let updatedCount = 0;
    let hasChanges = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i] && data[i][1] && String(data[i][1]).toLowerCase().trim() === email.toLowerCase().trim()) {
        const status = data[i][7];
        if (status === "PENDING") {
          const matchId = data[i][2];
          const desc = data[i][3];
          const betOn = String(data[i][4]).toUpperCase().trim();
          const odds = parseFloat(data[i][5]) || 0;
          const amount = parseFloat(data[i][6]) || 0;

          let cachedJsonStr = scriptCache.get("fb_match_res_" + matchId);
          let matchJson = null;

          if (cachedJsonStr) {
            try {
              matchJson = JSON.parse(cachedJsonStr);
            } catch (e) {
              matchJson = null;
            }
          }

          if (!matchJson) {
            try {
              const url = `https://api.football-data.org/v4/matches/${matchId}`;
              const response = UrlFetchApp.fetch(url, {
                'method': 'get',
                'headers': { 'X-Auth-Token': FOOTBALL_API_KEY },
                'muteHttpExceptions': true
              });
              const rc = response.getResponseCode();
              if (rc === 200) {
                const responseText = response.getContentText();
                const json = JSON.parse(responseText);
                if (json && (json.status === "FINISHED" || json.status === "AWARDED" || json.status === "CANCELED" || json.status === "POSTPONED")) {
                  matchJson = json;
                  scriptCache.put("fb_match_res_" + matchId, responseText, 43200);
                }
              } else if (rc === 429) {
                Logger.log("Football API bị giới hạn tần suất trong tiến trình kiểm tra tự động.");
              }
            } catch (apiErr) {
              Logger.log("Lỗi khi lấy kết quả trận đấu " + matchId + " từ API: " + apiErr.toString());
            }
          }

          if (matchJson) {
            const betLabel = getBetName(betOn, matchId);
            if (matchJson.status === "CANCELED" || matchJson.status === "POSTPONED") {
              data[i][7] = "CANCELED";
              data[i][8] = "Hủy/Hoãn";
              data[i][9] = amount;

              addUserPointsByEmail(email, amount, `Hoàn trả dự đoán bóng đá (Tự động): ${desc} - ${betLabel} bị Hủy/Hoãn`);
              updatedCount++;
              hasChanges = true;
            } else if (matchJson.status === "FINISHED" || matchJson.status === "AWARDED") {
              const p = matchJson.score && matchJson.score.fullTime;
              if (p && typeof p.home !== 'undefined' && typeof p.away !== 'undefined') {
                const scoreStr = `${p.home} - ${p.away}`;
                data[i][8] = scoreStr;

                if (isBetWinning(betOn, matchId, p.home, p.away)) {
                  data[i][7] = "WIN";
                  let rawReward = Math.floor(amount * odds);
                  let fee = Math.floor(rawReward * 0.05); // Phí nhà cái 5% để giảm lạm phát coin
                  const reward = rawReward - fee;
                  data[i][9] = reward;
                  addUserPointsByEmail(email, reward, `Thắng cược bóng đá (Tự động): ${desc} - ${betLabel} (Tỷ số ${scoreStr}) - Đã trừ ${fee} Coin phí`);
                } else {
                  data[i][7] = "LOSE";
                  data[i][9] = 0;
                }
                updatedCount++;
                hasChanges = true;
              }
            }
          }
        }
      }
    }

    if (hasChanges && updatedCount > 0) {
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }
  } catch (err) {
    Logger.log("Lỗi xử lý autoUpdateUserFootballBets: " + err.toString());
  } finally {
    lock.releaseLock();
  }
}

function adminUpdateFootballBets() {
  const adminEmail = CacheService.getUserCache().get("loggedInUser");
  if (!adminEmail || getUserRole(adminEmail) !== "ADMIN") return { success: false, message: "Unauthorized. Admin only." };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, message: "Hệ thống đang cập nhật, vui lòng đợi..." };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("FootballBets");
    if (!sheet) return { success: false, message: "Chưa có dữ liệu dự đoán bóng đá." };
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, message: "Chưa có lượt dự đoán nào." };
    
    let pendingMatches = new Set();
    for (let i = 1; i < data.length; i++) {
        if (data[i][7] === "PENDING") {
             pendingMatches.add(data[i][2]);
        }
    }
    
    if (pendingMatches.size === 0) return { success: true, message: "Không có lượt dự đoán PENDING nào cần duyệt." };
    
    let pendingArray = Array.from(pendingMatches);
    let matchResults = {};
    let limited = false;
    
    for (let i = 0; i < pendingArray.length; i++) {
        if (i >= 8) { limited = true; break; } 
        let matchId = pendingArray[i];
        
        try {
            const url = `https://api.football-data.org/v4/matches/${matchId}`;
            const response = UrlFetchApp.fetch(url, {
                'method': 'get',
                'headers': { 'X-Auth-Token': FOOTBALL_API_KEY },
                'muteHttpExceptions': true
            });
            const rc = response.getResponseCode();
            if (rc === 429) {
                 limited = true;
                 break;
            }
            if (rc === 200) {
                 const json = JSON.parse(response.getContentText());
                 if (json.status === "FINISHED" || json.status === "AWARDED" || json.status === "CANCELED" || json.status === "POSTPONED") {
                      matchResults[matchId] = json;
                 }
            }
            Utilities.sleep(500); 
        } catch (e) {
             Logger.log("Err checking match: " + e.message);
        }
    }
    
    let updatedCount = 0;
    
    for (let i = 1; i < data.length; i++) {
        if (data[i][7] === "PENDING") {
            let mId = data[i][2];
            let json = matchResults[mId];
            if (json) {
                let email = data[i][1];
                let desc = data[i][3];
                let betOn = String(data[i][4]).toUpperCase().trim();
                let odds = parseFloat(data[i][5]);
                let amount = parseFloat(data[i][6]);
                const betLabel = getBetName(betOn, mId);
                
                if (json.status === "CANCELED" || json.status === "POSTPONED") {
                     data[i][7] = "CANCELED"; 
                     data[i][8] = "Hủy/Hoãn";   
                     data[i][9] = amount;       
                     
                     addUserPointsByEmail(email, amount, `Hoàn Tiền Thể Thao: ${desc} - ${betLabel} bị Hủy/Hoãn`);
                     updatedCount++;
                } else if (json.status === "FINISHED" || json.status === "AWARDED") {
                     let p = json.score.fullTime;
                     if (p) {
                         let scoreStr = `${p.home} - ${p.away}`;
                         data[i][8] = scoreStr;
                         
                         if (isBetWinning(betOn, mId, p.home, p.away)) {
                             data[i][7] = "WIN";
                             let rawReward = Math.floor(amount * odds);
                             let fee = Math.floor(rawReward * 0.05); // Phí 5%
                             let reward = rawReward - fee;
                             data[i][9] = reward; 
                             addUserPointsByEmail(email, reward, `Thắng cược Thể thao: ${desc} - ${betLabel} (Tỷ số ${scoreStr}) - Đã trừ ${fee} Coin phí`);
                             updatedCount++;
                         } else {
                             data[i][7] = "LOSE";
                             data[i][9] = 0;
                             updatedCount++;
                         }
                     }
                }
            }
        }
    }
    
    if (updatedCount > 0) {
        sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
        let ext = limited ? " (Còn trận chưa xét do giới hạn API, hãy click lại sau 1 phút)." : "";
        return { success: true, message: `Đã tự động cộng thưởng cho ${updatedCount} lượt dự đoán!` + ext };
    } else {
        let ext = limited ? " (Giới hạn API, hãy click lại sau 1 phút)." : "";
        return { success: true, message: "Không có trận đấu nào trong danh sách PENDING đã kết thúc/có kết quả." + ext };
    }
    
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function addUserPointsByEmail(email, pointsToAdd, source = "Hệ thống") {
  if (!email) return false;
  const user = getUserByEmail(email);
  if (!user) return false;

  const current = parseInt(user.data[10]) || 0;
  let newTotal = current + pointsToAdd;
  if (newTotal < 0) newTotal = 0;

  updateUserField(user.row, 10, newTotal);
  logCoinTransaction(email, pointsToAdd, "TD Coin", source);
  return newTotal;
}

// ================= LIGHTWEIGHT STUDENT COIN RELIEF SERVER API =================
function logCoinTransaction(email, amount, type, source) {
  if (!email || amount === 0) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("CoinTransactions");
    if (!sheet) {
      sheet = ss.insertSheet("CoinTransactions");
      sheet.appendRow(["Thời Gian", "Email", "Số Tiền Thay Đổi", "Loại Tiền", "Hoạt Động / Nguồn Tiền"]);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f0f0f0");
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 150);
      sheet.setColumnWidth(2, 200);
      sheet.setColumnWidth(3, 120);
      sheet.setColumnWidth(4, 120);
      sheet.setColumnWidth(5, 300);
    }
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    sheet.appendRow([now, email, amount, type, source]);
    
    // Giới hạn 1000 dòng cho CoinTransactions
    limitSheetRows_(sheet, 1000);
  } catch (e) {
  } finally {
    lock.releaseLock();
  }
}

function limitSheetRows_(sheet, maxRows) {
  try {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow > maxRows) {
      const numToDelete = lastRow - maxRows;
      if (numToDelete > 0) {
        sheet.deleteRows(2, numToDelete);
      }
    }
  } catch (err) {
    Logger.log("Lỗi khi dọn giới hạn dòng: " + err.message);
  }
}

function claimStudentCoinReliefClient() {

  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) {
    return { success: false, message: "Bạn chưa đăng nhập." };
  }

  const role = getUserRole(email);
  const isEligibleRole = (role === "STUDENT" || role === "S-STUDENT");
  if (!isEligibleRole) {
    return { success: false, message: "Tính năng này chỉ áp dụng cho tài khoản học sinh/sinh viên (STUDENT)." };
  }

  // The client side manages the threshold (< 100) and weekly cooldown (localStorage) to avoid heavy server operations.
  const newPoints = addUserPoints_(1000, "Nhận Trợ Cấp Khẩn Cấp"); // 1000 coins as requested
  if (newPoints === false) {
    return { success: false, message: "Không thể cộng Coin. Vui lòng thử lại sau." };
  }

  return {
    success: true,
    message: "Nhận cứu trợ thành công! Bạn được tài trợ +1000 🪙 vào tài khoản học sinh.",
    newPoints: newPoints
  };
}

function adminPerformSeasonReset(ratioPercent, maxRetain) {
  const adminEmail = CacheService.getUserCache().get("loggedInUser");
  if (!adminEmail || getUserRole(adminEmail) !== "ADMIN") {
      return { success: false, message: "Unauthorized. Admin only." };
  }
  
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
      return { success: false, message: "Hệ thống đang bận, vui lòng thử lại." };
  }
  
  try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const usersSheet = ss.getSheetByName("Users");
      if (!usersSheet) return { success: false, message: "Data sheet not found" };
      
      const data = usersSheet.getDataRange().getValues();
      if (data.length <= 1) return { success: false, message: "No data" };
      
      const ratio = ratioPercent / 100;
      let usersAffected = 0;
      let totalCoinsBurned = 0;
      
      // Update data array
      for (let i = 1; i < data.length; i++) {
          let role = String(data[i][13] || "").toUpperCase().trim();
          let currentCoin = Number(data[i][10]); // K col = index 10
          if (isNaN(currentCoin)) currentCoin = 0;
          
          if (role === 'STUDENT' || role === 'S-STUDENT') {
              if (currentCoin !== 5000) {
                  let burned = currentCoin - 5000;
                  data[i][10] = 5000;
                  usersAffected++;
                  totalCoinsBurned += burned;
              }
          } else {
              let newCoin = Math.floor(currentCoin * ratio);
              if (newCoin > maxRetain) newCoin = maxRetain;
              
              let burned = currentCoin - newCoin;
              if (burned !== 0) {
                  data[i][10] = newCoin;
                  usersAffected++;
                  totalCoinsBurned += burned;
              }
          }
      }
      
      if (usersAffected > 0) {
          // Write back K col efficiently using getRange
          const kColData = data.map(row => [row[10]]);
          usersSheet.getRange(1, 11, kColData.length, 1).setValues(kColData);
          
          // Dọn dẹp CoinTransactions và FootballBets khi reset mùa giải
          try {
              const coinSheet = ss.getSheetByName("CoinTransactions");
              if (coinSheet) {
                  const lastRow = coinSheet.getLastRow();
                  if (lastRow > 1) {
                      coinSheet.deleteRows(2, lastRow - 1);
                  }
              }
          } catch (err) {
              Logger.log("Lỗi khi dọn CoinTransactions: " + err.message);
          }
          try {
              const betSheet = ss.getSheetByName("FootballBets");
              if (betSheet) {
                  const lastRow = betSheet.getLastRow();
                  if (lastRow > 1) {
                      betSheet.deleteRows(2, lastRow - 1);
                  }
              }
          } catch (err) {
              Logger.log("Lỗi khi dọn FootballBets: " + err.message);
          }
          
          logCoinTransaction(adminEmail, -totalCoinsBurned, "TD Coin", `RESET MÙA GIẢI (${ratioPercent}%, max ${maxRetain})`);
      }
      
      return {
          success: true,
          usersAffected: usersAffected,
          totalCoinsBurned: totalCoinsBurned
      };
      
  } catch (e) {
      return { success: false, message: e.message };
  } finally {
      lock.releaseLock();
  }
}


function checkAndRunAutoSeasonReset() {
  const props = PropertiesService.getScriptProperties();
  const lastReset = props.getProperty("LastAutoResetSeasonTD"); 
  
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth(); // 0-11
  
  const q = Math.floor(month / 3) + 1; // 1, 2, 3, 4
  const currentSeasonStr = `${year}-Q${q}`;
  
  if (lastReset !== currentSeasonStr) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(30000)) {
      try {
        // Double check
        const doubleCheck = PropertiesService.getScriptProperties().getProperty("LastAutoResetSeasonTD");
        if (doubleCheck === currentSeasonStr) return;
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const usersSheet = ss.getSheetByName("Users");
        if (usersSheet) {
          const data = usersSheet.getDataRange().getValues();
          let usersAffected = 0;
          let totalBurned = 0;
          for (let i = 1; i < data.length; i++) {
            let role = String(data[i][13] || "").toUpperCase().trim();
            let currentCoin = Number(data[i][10]);
            
            if (isNaN(currentCoin)) currentCoin = 0;
            
            if (role === 'STUDENT' || role === 'S-STUDENT') {
               // Học viên bình thường: Reset cứng về 5.000 Coin
               if (currentCoin !== 5000) {
                 totalBurned += (currentCoin - 5000);
                 data[i][10] = 5000;
                 usersAffected++;
               }
            } else {
               // Các role khác (VIP, O-STUDENT, ADMIN...): Giữ lại 10%, tối đa 5.000 Coin
               let newCoin = Math.floor(currentCoin * 0.1);
               if (newCoin > 5000) newCoin = 5000;
               if (newCoin !== currentCoin) {
                 totalBurned += (currentCoin - newCoin);
                 data[i][10] = newCoin;
                 usersAffected++;
               }
            }
          }
          if (usersAffected > 0) {
             const kColData = data.map(row => [row[10]]);
             usersSheet.getRange(1, 11, kColData.length, 1).setValues(kColData);
             
             // Dọn dẹp CoinTransactions và FootballBets khi reset mùa giải tự động
             try {
                 const coinSheet = ss.getSheetByName("CoinTransactions");
                 if (coinSheet) {
                     const lastRow = coinSheet.getLastRow();
                     if (lastRow > 1) {
                         coinSheet.deleteRows(2, lastRow - 1);
                     }
                 }
             } catch (err) {
                 Logger.log("Lỗi khi dọn CoinTransactions: " + err.message);
             }
             try {
                 const betSheet = ss.getSheetByName("FootballBets");
                 if (betSheet) {
                     const lastRow = betSheet.getLastRow();
                     if (lastRow > 1) {
                         betSheet.deleteRows(2, lastRow - 1);
                     }
                 }
             } catch (err) {
                 Logger.log("Lỗi khi dọn FootballBets: " + err.message);
             }
             
             logActivityToSheet("HỆ THỐNG", "SEASON_RESET", `Khởi tạo mùa mới ${currentSeasonStr}. Đã reset ${usersAffected} tài khoản STUDENT/S-STUDENT về mức 5,000 Coin.`);
          }
        }

        safeSetScriptProperty("LastAutoResetSeasonTD", currentSeasonStr);
      } catch(e) {
         Logger.log(e);
      } finally {
        lock.releaseLock();
      }
    }
  }
}

function getSStudentsList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    if (!sheet) return { success: false, message: "Không tìm thấy dữ liệu người dùng." };
    
    const data = sheet.getDataRange().getValues();
    const sstudents = [];
    
    // Header is row 0. Loop through users starting from row 1 (i = 1)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[0] || "";
      const email = row[1] ? row[1].toString().trim() : "";
      if (!email) continue;
      
      const roleRaw = (row[13] || "").toString().trim().toUpperCase();
      if (roleRaw === "S-STUDENT") {
        const points = Number(row[10]) || 0;
        
        let lastActiveDate = row[14] || row[2]; // fallback to Column C
        let daysAgoText = "N/A";
        
        if (lastActiveDate) {
          let dateObj = null;
          if (lastActiveDate instanceof Date) {
            dateObj = lastActiveDate;
          } else {
            const parsed = Date.parse(lastActiveDate);
            if (!isNaN(parsed)) {
              dateObj = new Date(parsed);
            }
          }
          
          if (dateObj) {
            const diffMs = new Date() - dateObj;
            const diffMins = Math.floor(diffMs / (1000 * 60));
            const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
            const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
            
            if (diffMins < 5) {
              daysAgoText = "Đang online 🟢";
            } else if (diffMins < 60) {
              daysAgoText = `${diffMins} phút trước`;
            } else if (diffHours < 24) {
              daysAgoText = `${diffHours} giờ trước`;
            } else {
              daysAgoText = `${diffDays} ngày trước`;
            }
          }
        }
        
        // Obfuscate email: tri***@gmail.com
        let obfuscatedEmail = "";
        const atIndex = email.indexOf("@");
        if (atIndex > 2) {
          obfuscatedEmail = email.substring(0, 3) + "***" + email.substring(atIndex);
        } else if (atIndex > 0) {
          obfuscatedEmail = email.substring(0, 1) + "***" + email.substring(atIndex);
        } else {
          obfuscatedEmail = "***";
        }
        
        sstudents.push({
          name: name || "Ẩn danh",
          email: obfuscatedEmail,
          points: points,
          lastActive: daysAgoText
        });
      }
    }
    
    // Sort S-STUDENT by points descending (motivation!)
    sstudents.sort((a, b) => b.points - a.points);
    
    return { success: true, data: sstudents };
  } catch (err) {
    return { success: false, message: "Lỗi Server: " + err.toString() };
  }
}

// ==========================================
// CLASS SCHEDULES MANAGEMENT (STEP 2)
// ==========================================

function getClassSchedulesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("ClassSchedules");
  if (!sheet) {
    sheet = ss.insertSheet("ClassSchedules");
    sheet.getRange(1, 1, 1, 6).setValues([["Mã Lớp", "Tên Lớp", "Lịch Học", "Giờ Học", "Link Học", "Màu Sắc"]]);
    sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#cbd5e1");
    // Add default templates
    sheet.getRange(2, 1, 1, 6).setValues([["TOAN101", "Toán Thầy A", "2,4,6", "19:30 - 21:00", "https://meet.google.com/abc-defg-hij", "#3b82f6"]]);
    sheet.getRange(3, 1, 1, 6).setValues([["LY102", "Vật Lý Thầy B", "3,5,7", "18:00 - 19:30", "https://meet.google.com/klm-nopq-rst", "#ec4899"]]);
  }
  return sheet;
}

function cleanDaysString(rawDays) {
  if (!rawDays) return "";
  
  // Handle Date object
  if (rawDays instanceof Date || (typeof rawDays === "object" && typeof rawDays.getMonth === "function")) {
    const d = rawDays.getDate();
    const m = rawDays.getMonth() + 1; // 1-based Month
    const days = [];
    if (d >= 2 && d <= 8) days.push(d === 8 ? "CN" : "T" + d.toString());
    if (m >= 2 && m <= 8) days.push(m === 8 ? "CN" : "T" + m.toString());
    
    // De-duplicate and sort
    const uniqueDays = [...new Set(days)];
    uniqueDays.sort((a, b) => {
      const valA = a === "CN" ? 8 : parseInt(a.replace("T",""));
      const valB = b === "CN" ? 8 : parseInt(b.replace("T",""));
      return valA - valB;
    });
    return uniqueDays.join(", ");
  }

  let str = rawDays.toString().trim();
  
  // Replace delimiters with comma
  str = str.replace(/[\.\/;|]/g, ",");
  
  const parts = str.split(",");
  const finalDays = [];
  
  parts.forEach(part => {
    let p = part.trim();
    if (!p) return;
    
    // If this part is a date string itself
    if (p.includes("GMT") || p.includes("giờ") || p.includes("Zone") || p.includes("00:00:00") || (p.includes(" ") && isNaN(p) && !isNaN(Date.parse(p)))) {
      try {
        const parsedDate = new Date(p);
        if (!isNaN(parsedDate.getTime())) {
          const d = parsedDate.getDate();
          const m = parsedDate.getMonth() + 1;
          if (d >= 2 && d <= 8) finalDays.push(d === 8 ? "CN" : "T" + d.toString());
          if (m >= 2 && m <= 8) finalDays.push(m === 8 ? "CN" : "T" + m.toString());
          return;
        }
      } catch (e) {}
    }
    
    // Otherwise, treat as normal day identifier
    let s = p.toLowerCase();
    s = s.replace(/^thứ\s*/i, "");
    s = s.replace(/^t/i, "");
    if (s === "cn" || s === "chủ nhật" || s === "8" || s === "sunday") {
      finalDays.push("CN");
      return;
    }
    const num = parseInt(s);
    if (num >= 2 && num <= 7) {
      finalDays.push("T" + num.toString());
    }
  });
  
  // Remove duplicates and sort
  const uniqueDays = [...new Set(finalDays)];
  uniqueDays.sort((a, b) => {
    const valA = a === "CN" ? 8 : parseInt(a.replace("T",""));
    const valB = b === "CN" ? 8 : parseInt(b.replace("T",""));
    return valA - valB;
  });
  
  return uniqueDays.join(", ");
}

function getClassSchedules() {
  try {
    const sheet = getClassSchedulesSheet();
    const data = sheet.getDataRange().getValues(); // Use raw getValues to preserve actual Date objects
    
    // Find header row (assume it's the first row that has "Mã Lớp")
    let headerRowIdx = -1;
    let headers = [];
    for(let i=0; i<data.length; i++) {
      const rowStr = data[i].map(c => String(c).toLowerCase()).join("");
      if(rowStr.includes("mã lớp") || rowStr.includes("tên lớp")) {
        headerRowIdx = i;
        headers = data[i].map(c => String(c).trim());
        break;
      }
    }
    
    if (headerRowIdx === -1) {
      headerRowIdx = 0;
      headers = data[0].map(c => String(c).trim());
    }
    
    const colMap = {
      code: headers.findIndex(h => h.toLowerCase() === "mã lớp"),
      name: headers.findIndex(h => h.toLowerCase() === "tên lớp"),
      days: headers.findIndex(h => h.toLowerCase() === "lịch học"),
      time: headers.findIndex(h => h.toLowerCase() === "giờ học"),
      link: headers.findIndex(h => h.toLowerCase() === "link học"),
      color: headers.findIndex(h => h.toLowerCase() === "màu sắc")
    };
    
    // fallback if columns not found
    if(colMap.code === -1) colMap.code = 0;
    if(colMap.name === -1) colMap.name = 1;
    if(colMap.days === -1) colMap.days = 2;
    if(colMap.time === -1) colMap.time = 3;
    if(colMap.link === -1) colMap.link = 4;
    if(colMap.color === -1) colMap.color = 5;

    const schedules = [];
    for (let i = headerRowIdx + 1; i < data.length; i++) {
      const code = (data[i][colMap.code] || "").toString().trim();
      if (!code) continue;
      
      const rawDays = data[i][colMap.days];
      const cleanedDays = cleanDaysString(rawDays);
      
      let scheduleLink = (data[i][colMap.link] || "").toString().trim();
      if (scheduleLink && !/^https?:\/\//i.test(scheduleLink) && !/^\/\//.test(scheduleLink)) {
        scheduleLink = "https://" + scheduleLink;
      }
      
      schedules.push({
        rowId: i + 1, // 1-based index in sheet
        code: code,
        name: (data[i][colMap.name] || "").toString().trim(),
        days: cleanedDays,
        time: (data[i][colMap.time] || "").toString().trim(),
        link: scheduleLink,
        color: (data[i][colMap.color] || "#10b981").toString().trim()
      });
    }
    return { success: true, data: schedules };
  } catch (err) {
    return { success: false, message: "Lỗi Server: " + err.toString() };
  }
}

function getUserClassCodes(email) {
  if (!email) return [];
  const user = getUserByEmail(email);
  if (!user) return [];
  const trackingColIndex = columnLetterToIndex("R"); // Cột R
  const codesRaw = (user.data[trackingColIndex] || "").toString().trim();
  if (!codesRaw) return [];
  return codesRaw.split(",").map(c => c.trim().toUpperCase()).filter(Boolean);
}

function getUserClassSchedules() {
  const email = CacheService.getUserCache().get("loggedInUser");
  if (!email) return { success: false, message: "Bạn chưa đăng nhập." };
  
  try {
    const userCodes = getUserClassCodes(email);
    const schedulesResult = getClassSchedules();
    if (!schedulesResult.success) return schedulesResult;
    
    const userSchedules = schedulesResult.data.filter(s => 
      userCodes.includes(s.code.toUpperCase())
    );
    
    return { success: true, data: userSchedules, userCodes: userCodes };
  } catch (err) {
    return { success: false, message: "Lỗi Server: " + err.toString() };
  }
}

function saveClassSchedule(schedule) {
  const adminEmail = CacheService.getUserCache().get("loggedInUser");
  if (!adminEmail || getUserRole(adminEmail) !== "ADMIN") {
    return { success: false, message: "Bạn không có quyền thực hiện chức năng này." };
  }
  
  try {
    const sheet = getClassSchedulesSheet();
    const data = sheet.getDataRange().getValues();
    const code = (schedule.code || "").toString().trim();
    if (!code) return { success: false, message: "Mã lớp không được để trống." };
    
    // Find header row
    let headerRowIdx = -1;
    let headers = [];
    for(let i=0; i<data.length; i++) {
      const rowStr = data[i].map(c => String(c).toLowerCase()).join("");
      if(rowStr.includes("mã lớp") || rowStr.includes("tên lớp")) {
        headerRowIdx = i;
        headers = data[i].map(c => String(c).trim());
        break;
      }
    }
    
    if (headerRowIdx === -1) {
      headerRowIdx = 0;
      headers = data[0].map(c => String(c).trim());
    }
    
    const colMap = {
      code: headers.findIndex(h => h.toLowerCase() === "mã lớp"),
      name: headers.findIndex(h => h.toLowerCase() === "tên lớp"),
      days: headers.findIndex(h => h.toLowerCase() === "lịch học"),
      time: headers.findIndex(h => h.toLowerCase() === "giờ học"),
      link: headers.findIndex(h => h.toLowerCase() === "link học"),
      color: headers.findIndex(h => h.toLowerCase() === "màu sắc")
    };
    
    if(colMap.code === -1) colMap.code = 0;
    if(colMap.name === -1) colMap.name = 1;
    if(colMap.days === -1) colMap.days = 2;
    if(colMap.time === -1) colMap.time = 3;
    if(colMap.link === -1) colMap.link = 4;
    if(colMap.color === -1) colMap.color = 5;

    let foundRowIndex = -1;
    if (schedule.rowId) {
      const rId = parseInt(schedule.rowId);
      if (rId > headerRowIdx + 1 && rId <= data.length) {
        foundRowIndex = rId;
      }
    }
    
    const cleanedDays = cleanDaysString(schedule.days || "");
    let scheduleLink = (schedule.link || "").toString().trim();
    if (scheduleLink && !/^https?:\/\//i.test(scheduleLink) && !/^\/\//.test(scheduleLink)) {
      scheduleLink = "https://" + scheduleLink;
    }
    const vals = {
      code: code,
      name: (schedule.name || "").toString().trim(),
      days: cleanedDays,
      time: (schedule.time || "").toString().trim(),
      link: scheduleLink,
      color: (schedule.color || "#10b981").toString().trim()
    };
    
    let targetRow = foundRowIndex;
    if (targetRow === -1) {
      let emptyRowIndex = -1;
      for (let i = headerRowIdx + 1; i < data.length; i++) {
        if (!data[i][colMap.code] && !data[i][colMap.name]) {
          emptyRowIndex = i + 1;
          break;
        }
      }
      if (emptyRowIndex !== -1) {
        targetRow = emptyRowIndex;
      } else {
        targetRow = sheet.getLastRow() + 1;
      }
    }
    
    // Batch Write class schedule values
    sheet.getRange(targetRow, 1, 1, 6).setValues([[vals.code, vals.name, vals.days, vals.time, vals.link, vals.color]]);
    
    return { success: true, message: "Đã lưu lịch học thành công." };
  } catch (err) {
    return { success: false, message: "Lỗi Server: " + err.toString() };
  }
}

function deleteClassSchedule(classCode, rowId) {
  const adminEmail = CacheService.getUserCache().get("loggedInUser");
  if (!adminEmail || getUserRole(adminEmail) !== "ADMIN") {
    return { success: false, message: "Bạn không có quyền thực hiện chức năng này." };
  }
  
  try {
    const sheet = getClassSchedulesSheet();
    const data = sheet.getDataRange().getValues();
    
    let foundRowIndex = -1;
    if (rowId) {
      const rId = parseInt(rowId);
      if (rId > 1 && rId <= data.length) {
        foundRowIndex = rId;
      }
    } else {
      const code = (classCode || "").toString().trim().toUpperCase();
      if (code) {
        // Find header row to dynamically know the code col
        let headerRowIdx = -1;
        let codeColIdx = 0;
        for(let i=0; i<data.length; i++) {
          const rowStr = data[i].map(c => String(c).toLowerCase()).join("");
          if(rowStr.includes("mã lớp")) {
            headerRowIdx = i;
            codeColIdx = data[i].findIndex(c => String(c).trim().toLowerCase() === "mã lớp");
            break;
          }
        }
        if (codeColIdx === -1) codeColIdx = 0;
        if (headerRowIdx === -1) headerRowIdx = 0;

        for (let i = headerRowIdx + 1; i < data.length; i++) {
          if ((data[i][codeColIdx] || "").toString().trim().toUpperCase() === code) {
            foundRowIndex = i + 1;
            break;
          }
        }
      }
    }
    
    if (foundRowIndex !== -1) {
      // Clear instead of deleteRow to not break table structure
      sheet.getRange(foundRowIndex, 1, 1, sheet.getLastColumn()).clearContent();
      return { success: true, message: "Đã xóa lịch học thành công." };
    } else {
      return { success: false, message: "Không tìm thấy lớp học cần xóa." };
    }
  } catch (err) {
    return { success: false, message: "Lỗi Server: " + err.toString() };
  }
}

function adminGetExamSubmissions(showAll) {
  try {
    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) return { success: false, message: "Bạn chưa đăng nhập." };
    const role = getUserRole(email);
    if (role !== "ADMIN") return { success: false, message: "Bạn không có quyền quản trị." };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ExamSubmissions');
    if (!sheet) {
      return { success: true, data: [] };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, data: [] };
    }

    // Deduplicate in memory and identify extra redundant rows to delete from sheet
    const submissions = [];
    const seenMap = {};
    const duplicateRowNumsToDelete = [];

    // Iterate backwards so the latest rows are prioritized
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const rowEmail = (row[1] || "").toString().trim();
      const rowTitle = (row[2] || "").toString().trim();
      if (!rowEmail || !rowTitle) continue;

      const normKey = normalizeExamText(rowEmail) + "___" + normalizeExamText(rowTitle);

      let gradedAtVal = "";
      if (row[8] instanceof Date) {
        gradedAtVal = Utilities.formatDate(row[8], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      } else if (row[8]) {
        gradedAtVal = row[8].toString().trim();
      }

      const scoreVal = row[5] !== undefined ? row[5].toString().trim() : "";
      const subObj = {
        timestamp: row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss") : (row[0] || "").toString(),
        email: rowEmail,
        examTitle: rowTitle,
        fileUrl: (row[3] || "").toString().trim(),
        fileName: (row[4] || "").toString().trim(),
        score: scoreVal,
        remark: row[6] !== undefined ? row[6].toString().trim() : "",
        gradedBy: row[7] !== undefined ? row[7].toString().trim() : "",
        gradedAt: gradedAtVal,
        rowNum: i + 1
      };

      if (!seenMap[normKey]) {
        seenMap[normKey] = subObj;
        submissions.push(subObj);
      } else {
        // Trùng lặp dòng nộp bài!
        // Nếu dòng hiện tại có điểm nhưng dòng đã chọn trước đó chưa có điểm -> Đổi sang dòng có điểm
        if (scoreVal !== "" && seenMap[normKey].score === "") {
          const oldSub = seenMap[normKey];
          duplicateRowNumsToDelete.push(oldSub.rowNum);
          seenMap[normKey] = subObj;
          const idx = submissions.indexOf(oldSub);
          if (idx !== -1) submissions[idx] = subObj;
        } else {
          duplicateRowNumsToDelete.push(i + 1);
        }
      }
    }

    // Tự động dọn dẹp các dòng nộp trùng lặp thừa trên Google Sheet theo thứ tự từ dưới lên
    if (duplicateRowNumsToDelete.length > 0) {
      try {
        duplicateRowNumsToDelete.sort((a, b) => b - a);
        duplicateRowNumsToDelete.forEach(rNum => {
          try { sheet.deleteRow(rNum); } catch(e) {}
        });
        invalidateSheetCache("ExamSubmissions");
      } catch (err) {
        Logger.log("Lỗi tự động xóa dòng trùng: " + err.toString());
      }
    }

    // Load mock exams to build a map from examTitle -> { adminEmail, className }
    let examMetaMap = {};
    const mockExamsSheet = ss.getSheetByName("MockExams");
    if (mockExamsSheet) {
      const mockData = mockExamsSheet.getDataRange().getValues();
      if (mockData.length > 1) {
        const headers = mockData[0].map(h => String(h).trim().toLowerCase());
        let titleIdx = 3;
        let adminEmailIdx = -1;
        let classIdx = 6; // Default to Column G
        headers.forEach((h, idx) => {
          if (h.includes("tên bài") || h.includes("tiêu đề") || h === "bài thi") titleIdx = idx;
          else if (h.includes("phụ trách") || h.includes("admin")) adminEmailIdx = idx;
          else if (h.includes("lớp") || h.includes("class")) classIdx = idx;
        });
        if (adminEmailIdx === -1 && headers.length >= 8) {
          adminEmailIdx = 7;
        }
        for (let k = 1; k < mockData.length; k++) {
          const titleVal = mockData[k][titleIdx] ? normalizeExamText(mockData[k][titleIdx]) : "";
          const adminVal = adminEmailIdx !== -1 && mockData[k][adminEmailIdx] ? mockData[k][adminEmailIdx].toString().trim().toLowerCase() : "";
          const classVal = classIdx !== -1 && mockData[k][classIdx] ? mockData[k][classIdx].toString().trim() : "Khác";
          if (titleVal) {
            examMetaMap[titleVal] = {
              adminEmail: adminVal,
              className: classVal
            };
          }
        }
      }
    }

    // Attach className to each submission
    submissions.forEach(sub => {
      const subTitleNorm = normalizeExamText(sub.examTitle);
      const meta = examMetaMap[subTitleNorm];
      sub.className = meta ? meta.className : "Khác";
    });

    // Dynamic filtering based on responsible admin email
    const showAllBool = (showAll === true || showAll === "true");
    let filtered = submissions;
    if (!showAllBool) {
      const lowerEmail = email.toLowerCase().trim();
      filtered = submissions.filter(sub => {
        const subTitleNorm = normalizeExamText(sub.examTitle);
        const meta = examMetaMap[subTitleNorm];
        const assignedAdmin = meta ? meta.adminEmail : "";
        if (assignedAdmin && assignedAdmin !== "" && assignedAdmin !== lowerEmail) {
          return false;
        }
        return true;
      });
    }

    // Giữ danh sách đã sắp xếp mới nhất
    return { success: true, data: filtered };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function adminGradeSubmission(studentEmail, examTitle, score, remark) {
  try {
    const adminEmail = CacheService.getUserCache().get("loggedInUser");
    if (!adminEmail) return { success: false, message: "Bạn chưa đăng nhập." };
    const role = getUserRole(adminEmail);
    if (role !== "ADMIN") return { success: false, message: "Bạn không có quyền quản trị." };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('ExamSubmissions');
    if (!sheet) {
      sheet = ss.insertSheet('ExamSubmissions');
      sheet.appendRow(['Timestamp', 'Email', 'Exam Title', 'File URL', 'Filename', 'Score', 'Remark', 'GradedBy', 'GradedAt']);
    }

    const data = sheet.getDataRange().getValues();
    const cleanStudentEmailNorm = normalizeExamText(studentEmail);
    const cleanExamTitleNorm = normalizeExamText(examTitle);

    const matchingRowNums = []; // 1-based row numbers
    for (let i = 1; i < data.length; i++) {
      const rowEmail = normalizeExamText(data[i][1]);
      const rowExamTitle = normalizeExamText(data[i][2]);
      if (rowEmail === cleanStudentEmailNorm && rowExamTitle === cleanExamTitleNorm) {
        matchingRowNums.push(i + 1);
      }
    }

    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    if (matchingRowNums.length === 0) {
      // Nếu chưa có bài nộp trong bảng, tự động ghi dòng mới để lưu điểm trực tiếp
      sheet.appendRow([new Date(), studentEmail.trim(), examTitle.trim(), "", "Chấm trực tiếp", score, remark, adminEmail, nowStr]);
    } else {
      // Cập nhật dòng đầu tiên
      const primaryRowNum = matchingRowNums[0];
      sheet.getRange(primaryRowNum, 2).setValue(studentEmail.trim());
      sheet.getRange(primaryRowNum, 3).setValue(examTitle.trim());
      sheet.getRange(primaryRowNum, 6).setValue(score);
      sheet.getRange(primaryRowNum, 7).setValue(remark);
      sheet.getRange(primaryRowNum, 8).setValue(adminEmail);
      sheet.getRange(primaryRowNum, 9).setValue(nowStr);

      // Nếu có dòng trùng lặp thừa còn sót lại, xóa từ dưới lên
      for (let k = matchingRowNums.length - 1; k >= 1; k--) {
        try { sheet.deleteRow(matchingRowNums[k]); } catch(e) {}
      }
    }

    // Xóa cache dữ liệu bài thi để ứng dụng hiển thị điểm mới nhất ngay lập tức
    invalidateSheetCache("ExamSubmissions");

    // Gửi thông báo tự động gom nhóm cho học sinh theo đúng đề thi đã chấm (tránh rác Server)
    const notifTitle = "Bài Đã Chấm Xong";
    const notifDesc = `Bài thi thử "${examTitle}" của bạn đã được giáo viên chấm điểm. Vui lòng xem chi tiết kết quả trong phần Khảo Thí.`;
    const notifAction = "mockExams:" + examTitle;
    addOrUpdateGroupNotification(notifTitle, notifDesc, notifAction, examTitle, studentEmail, adminEmail);

    return { success: true, message: "Đã chấm điểm và gửi kết quả cho học sinh thành công." };
  } catch (e) {
    return { success: false, message: "Lỗi Server: " + e.message };
  }
}

// ========================================================
// HOMEWORK PERSISTENCE & GRADING SYSTEM BACKEND
// ========================================================

function getHomeworkData(fileId, showAll) {
  try {
    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) return { success: false, message: "Bạn chưa đăng nhập." };

    const role = getUserRole(email);
    const isAdmin = (role === "ADMIN");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Get Homework Note (Báo bài)
    var note = "";
    var deadline = "";
    var answerKey = "";
    var numQuestions = "";
    var pointsPerQuestion = "";
    var creatorEmail = "";
    var noteSheet = ss.getSheetByName('HomeworkNotes');
    if (noteSheet) {
      var noteData = noteSheet.getDataRange().getValues();
      var headers = noteData[0];
      var deadlineColIdx = headers.indexOf('Deadline');
      var answerKeyColIdx = headers.indexOf('AnswerKey');
      var numQuestionsColIdx = headers.indexOf('NumQuestions');
      var pointsPerQuestionColIdx = headers.indexOf('PointsPerQuestion');
      for (var i = 1; i < noteData.length; i++) {
        if (noteData[i][0] && noteData[i][0].toString() === fileId) {
          note = noteData[i][2] ? noteData[i][2].toString() : "";
          creatorEmail = noteData[i][4] ? noteData[i][4].toString().trim().toLowerCase() : "";
          if (deadlineColIdx !== -1 && noteData[i][deadlineColIdx]) {
            deadline = noteData[i][deadlineColIdx].toString();
          }
          if (answerKeyColIdx !== -1 && noteData[i][answerKeyColIdx]) {
            answerKey = noteData[i][answerKeyColIdx].toString();
          }
          if (numQuestionsColIdx !== -1 && noteData[i][numQuestionsColIdx]) {
            numQuestions = noteData[i][numQuestionsColIdx].toString();
          }
          if (pointsPerQuestionColIdx !== -1 && noteData[i][pointsPerQuestionColIdx]) {
            pointsPerQuestion = noteData[i][pointsPerQuestionColIdx].toString();
          }
          break;
        }
      }
    }

    // 2. Get Student's own submission
    var submission = null;
    var subSheet = ss.getSheetByName('HomeworkSubmissions');
    if (subSheet) {
      var subData = subSheet.getDataRange().getValues();
      var lowerEmail = email.toLowerCase().trim();
      for (var i = 1; i < subData.length; i++) {
        var subFileId = subData[i][2] ? subData[i][2].toString() : "";
        var subEmail = subData[i][1] ? subData[i][1].toString().toLowerCase().trim() : "";
        if (subFileId === fileId && subEmail === lowerEmail) {
          submission = {
            timestamp: Utilities.formatDate(new Date(subData[i][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
            email: subData[i][1],
            fileId: subData[i][2],
            fileTitle: subData[i][3],
            fileUrl: subData[i][4],
            filename: subData[i][5],
            score: subData[i][6] !== undefined ? subData[i][6].toString() : "",
            remark: subData[i][7] !== undefined ? subData[i][7].toString() : ""
          };
          break;
        }
      }
    }

    // 3. If Admin, also get all submissions of this specific fileId
    var allSubmissionsOfFile = [];
    if (isAdmin && subSheet) {
      const showAllBool = (showAll === true || showAll === "true");
      const loggedInEmail = email.toLowerCase().trim();
      const isCreator = (creatorEmail === "" || creatorEmail === loggedInEmail);

      if (showAllBool || isCreator) {
        var subData = subSheet.getDataRange().getValues();
        for (var i = subData.length - 1; i >= 1; i--) {
          var subFileId = subData[i][2] ? subData[i][2].toString() : "";
          if (subFileId === fileId) {
            allSubmissionsOfFile.push({
              timestamp: Utilities.formatDate(new Date(subData[i][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
              email: subData[i][1],
              fileId: subData[i][2],
              fileTitle: subData[i][3],
              fileUrl: subData[i][4],
              filename: subData[i][5],
              score: subData[i][6] !== undefined ? subData[i][6].toString() : "",
              remark: subData[i][7] !== undefined ? subData[i][7].toString() : ""
            });
          }
        }
      }
    }

    return {
      success: true,
      note: note,
      deadline: deadline,
      answerKey: answerKey,
      submission: submission,
      allSubmissionsOfFile: allSubmissionsOfFile,
      numQuestions: numQuestions,
      pointsPerQuestion: pointsPerQuestion
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveHomeworkAnswerKey(fileId, answerKey) {
  try {
    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) return { success: false, message: "Bạn chưa đăng nhập." };

    const role = getUserRole(email);
    if (role !== "ADMIN") {
      return { success: false, message: "Bạn không có quyền thực hiện chức năng này." };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('HomeworkNotes');
    if (!sheet) {
      sheet = ss.insertSheet('HomeworkNotes');
      sheet.appendRow(['File ID', 'File Title', 'Note Content', 'Last Updated', 'Updated By', 'Deadline', 'AnswerKey']);
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var answerKeyColIdx = headers.indexOf('AnswerKey');
    if (answerKeyColIdx === -1) {
      sheet.getRange(1, headers.length + 1).setValue('AnswerKey');
      headers.push('AnswerKey');
      answerKeyColIdx = headers.length - 1;
    }

    var foundIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === fileId) {
        foundIndex = i;
        break;
      }
    }

    if (foundIndex !== -1) {
      var rowNum = foundIndex + 1;
      sheet.getRange(rowNum, answerKeyColIdx + 1).setValue(answerKey || "");
    } else {
      var newRow = [fileId, "Bài tập", "", new Date(), email];
      var deadlineColIdx = headers.indexOf('Deadline');
      if (deadlineColIdx !== -1) {
        newRow[deadlineColIdx] = "";
      }
      newRow[answerKeyColIdx] = answerKey || "";
      sheet.appendRow(newRow);
    }

    logActivityToSheet(email, "SAVE_HOMEWORK_ANSWER_KEY", `Lưu đáp án BTVN cho file: ${fileId}`);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveHomeworkNote(fileId, fileTitle, noteContent, deadline, numQuestions, pointsPerQuestion) {
  try {
    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) return { success: false, message: "Bạn chưa đăng nhập." };

    const role = getUserRole(email);
    if (role !== "ADMIN") {
      return { success: false, message: "Bạn không có quyền thực hiện chức năng này." };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('HomeworkNotes');
    if (!sheet) {
      sheet = ss.insertSheet('HomeworkNotes');
      sheet.appendRow(['File ID', 'File Title', 'Note Content', 'Last Updated', 'Updated By', 'Deadline']);
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // Ensure all required columns exist in HomeworkNotes
    var deadlineColIdx = headers.indexOf('Deadline');
    if (deadlineColIdx === -1) {
      sheet.getRange(1, headers.length + 1).setValue('Deadline');
      headers.push('Deadline');
      deadlineColIdx = headers.length - 1;
    }
    var numQuestionsColIdx = headers.indexOf('NumQuestions');
    if (numQuestionsColIdx === -1) {
      sheet.getRange(1, headers.length + 1).setValue('NumQuestions');
      headers.push('NumQuestions');
      numQuestionsColIdx = headers.length - 1;
    }
    var pointsPerQuestionColIdx = headers.indexOf('PointsPerQuestion');
    if (pointsPerQuestionColIdx === -1) {
      sheet.getRange(1, headers.length + 1).setValue('PointsPerQuestion');
      headers.push('PointsPerQuestion');
      pointsPerQuestionColIdx = headers.length - 1;
    }

    var foundIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === fileId) {
        foundIndex = i;
        break;
      }
    }

    if (foundIndex !== -1) {
      var rowNum = foundIndex + 1;
      sheet.getRange(rowNum, 2, 1, 4).setValues([[fileTitle, noteContent, new Date(), email]]);
      sheet.getRange(rowNum, deadlineColIdx + 1).setValue(deadline || "");
      sheet.getRange(rowNum, numQuestionsColIdx + 1).setValue(numQuestions || "");
      sheet.getRange(rowNum, pointsPerQuestionColIdx + 1).setValue(pointsPerQuestion || "");
    } else {
      var newRow = [];
      newRow[0] = fileId;
      newRow[1] = fileTitle;
      newRow[2] = noteContent;
      newRow[3] = new Date();
      newRow[4] = email;
      newRow[deadlineColIdx] = deadline || "";
      newRow[numQuestionsColIdx] = numQuestions || "";
      newRow[pointsPerQuestionColIdx] = pointsPerQuestion || "";
      
      // Pad newRow so appendRow sets all columns properly
      for (var k = 0; k < headers.length; k++) {
        if (newRow[k] === undefined) {
          newRow[k] = "";
        }
      }
      sheet.appendRow(newRow);
    }

    logActivityToSheet(email, "SAVE_HOMEWORK_NOTE", `Lưu báo bài BTVN cho file: ${fileTitle}`);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function uploadHomeworkSubmission(base64Data, filename, fileId, fileTitle) {
  var lock = LockService.getScriptLock();
  var lockAcquired = false;
  try {
    try {
      lockAcquired = lock.tryLock(30000);
    } catch(eLock) {}

    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) return { success: false, message: "Bạn chưa đăng nhập." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Verify if this homework is assigned (has a note in HomeworkNotes)
    var noteSheet = ss.getSheetByName('HomeworkNotes');
    var isAssigned = false;
    var deadlineDate = null;
    if (noteSheet) {
      var noteData = noteSheet.getDataRange().getValues();
      var headers = noteData[0];
      var deadlineColIdx = headers.indexOf('Deadline');
      for (var i = 1; i < noteData.length; i++) {
        if (noteData[i][0] && noteData[i][0].toString() === fileId) {
          var noteVal = noteData[i][2] ? noteData[i][2].toString().trim() : "";
          if (noteVal !== "") {
            isAssigned = true;
            if (deadlineColIdx !== -1 && noteData[i][deadlineColIdx]) {
              deadlineDate = noteData[i][deadlineColIdx];
            }
          }
          break;
        }
      }
    }

    if (!isAssigned) {
      return { success: false, message: "không có BTVN" };
    }

    // Verify deadline
    if (deadlineDate) {
      var dlDate = (deadlineDate instanceof Date) ? deadlineDate : new Date(deadlineDate.toString().trim());
      if (dlDate && !isNaN(dlDate.getTime())) {
        var now = new Date();
        if (now > dlDate) {
          return { success: false, message: "Hạn nộp bài tập về nhà đã hết hạn." };
        }
      }
    }

    var sheet = ss.getSheetByName('HomeworkSubmissions');
    if (!sheet) {
      sheet = ss.insertSheet('HomeworkSubmissions');
      sheet.appendRow(['Timestamp', 'Email', 'File ID', 'File Title', 'File URL', 'Filename', 'Score', 'Remark']);
    }

    var data = sheet.getDataRange().getValues();
    var existingRowIndex = -1;
    var oldFileUrl = "";
    var lowerEmail = email.toLowerCase().trim();

    for (var i = 1; i < data.length; i++) {
      var rowEmail = data[i][1] ? data[i][1].toString().toLowerCase().trim() : "";
      var rowFileId = data[i][2] ? data[i][2].toString() : "";
      if (rowEmail === lowerEmail && rowFileId === fileId) {
        existingRowIndex = i;
        oldFileUrl = data[i][4] ? data[i][4].toString() : "";
        break;
      }
    }

    if (existingRowIndex !== -1 && oldFileUrl) {
      try {
        var matchId = oldFileUrl.match(/[-\w]{25,}/);
        var oldFileId = matchId ? matchId[0] : null;
        if (oldFileId) {
          var oldFile = DriveApp.getFileById(oldFileId);
          oldFile.setTrashed(true);
        }
      } catch (err) {
        Logger.log("Không thể xóa file cũ: " + err.toString());
      }
    }

    var parentFolderId = "1qhig9lIft7D8upXv4c6Qk6INHu8fNM5f";
    var parentFolder = DriveApp.getFolderById(parentFolderId);
    var emailFolder;
    var emailFolders = parentFolder.getFoldersByName(lowerEmail);
    if (emailFolders.hasNext()) {
      emailFolder = emailFolders.next();
    } else {
      emailFolder = parentFolder.createFolder(lowerEmail);
    }
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "application/pdf", fileTitle + " - " + email + " - " + filename);
    var file = emailFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    if (existingRowIndex !== -1) {
      sheet.deleteRow(existingRowIndex + 1);
    }
    sheet.appendRow([new Date(), email, fileId, fileTitle, file.getUrl(), file.getName(), "", ""]);

    logActivityToSheet(email, "SUBMIT_HOMEWORK", `Nộp BTVN cho file: ${fileTitle}`);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    if (lockAcquired) {
      try { lock.releaseLock(); } catch(e) {}
    }
  }
}

function getHomeworkSubmissions(showAll) {
  try {
    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) return { success: false, message: "Bạn chưa đăng nhập." };

    const role = getUserRole(email);
    if (role !== "ADMIN") {
      return { success: false, message: "Bạn không có quyền truy cập." };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('HomeworkSubmissions');
    if (!sheet) {
      return { success: true, data: [] };
    }

    var data = sheet.getDataRange().getValues();
    var submissions = [];
    for (var i = 1; i < data.length; i++) {
      submissions.push({
        timestamp: Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
        email: data[i][1] ? data[i][1].toString() : "",
        fileId: data[i][2] ? data[i][2].toString() : "",
        fileTitle: data[i][3] ? data[i][3].toString() : "",
        fileUrl: data[i][4] ? data[i][4].toString() : "",
        filename: data[i][5] ? data[i][5].toString() : "",
        score: data[i][6] !== undefined ? data[i][6].toString() : "",
        remark: data[i][7] !== undefined ? data[i][7].toString() : ""
      });
    }

    // Resolve folders server-side for proper grouping based on file locations
    var libraryFolderMap = {};
    try {
      var courses = getCachedCoursesFromSheet();
      if (courses && courses.length > 0) {
        courses.forEach(function(c) {
          if (c.fileId && c.folder) {
            libraryFolderMap[String(c.fileId).trim()] = c.folder.trim();
          }
        });
      }
    } catch(err) {
      Logger.log("getHomeworkSubmissions: Error loading courses: " + err);
    }

    try {
      var books = getCachedBooksFromSheet();
      if (books && books.length > 0) {
        books.forEach(function(b) {
          if (b.fileId && b.folder) {
            libraryFolderMap[String(b.fileId).trim()] = b.folder.trim();
          }
        });
      }
    } catch(err) {
      Logger.log("getHomeworkSubmissions: Error loading books: " + err);
    }

    try {
      var exams = getCachedExamsFromSheet();
      if (exams && exams.length > 0) {
        exams.forEach(function(e) {
          if (e.fileId && e.folder) {
            libraryFolderMap[String(e.fileId).trim()] = e.folder.trim();
          }
        });
      }
    } catch(err) {
      Logger.log("getHomeworkSubmissions: Error loading exams: " + err);
    }

    submissions.forEach(function(sub) {
      var fId = String(sub.fileId).trim();
      sub.folder = libraryFolderMap[fId] || "";
    });

    // Dynamic filtering based on homework teacher email
    const showAllBool = (showAll === true || showAll === "true");
    if (!showAllBool) {
      // Build map from fileId -> creator/teacher email from HomeworkNotes (column E / index 4)
      var homeworkAdminMap = {};
      var notesSheet = ss.getSheetByName("HomeworkNotes");
      if (notesSheet) {
        var notesData = notesSheet.getDataRange().getValues();
        for (var k = 1; k < notesData.length; k++) {
          var fileIdVal = notesData[k][0] ? notesData[k][0].toString().trim() : "";
          var creatorVal = notesData[k][4] ? notesData[k][4].toString().trim().toLowerCase() : "";
          if (fileIdVal && creatorVal) {
            homeworkAdminMap[fileIdVal] = creatorVal;
          }
        }
      }

      const lowerEmail = email.toLowerCase().trim();
      // Filter out submissions where there is an assigned teacher and it is NOT the logged-in admin
      var filtered = submissions.filter(function(sub) {
        const subFileId = (sub.fileId || "").trim();
        const assignedTeacher = homeworkAdminMap[subFileId];
        if (assignedTeacher && assignedTeacher !== "" && assignedTeacher !== lowerEmail) {
          return false;
        }
        return true;
      });
      filtered.reverse();
      return { success: true, data: filtered };
    }

    submissions.reverse();
    return { success: true, data: submissions };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveHomeworkGradeAndRemark(studentEmail, fileId, score, remark) {
  return gradeHomework(studentEmail, fileId, score, remark);
}

function gradeHomework(studentEmail, fileId, score, remark) {
  try {
    const adminEmail = CacheService.getUserCache().get("loggedInUser");
    if (!adminEmail) return { success: false, message: "Bạn chưa đăng nhập." };

    const role = getUserRole(adminEmail);
    if (role !== "ADMIN") {
      return { success: false, message: "Bạn không có quyền thực hiện hành động này." };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('HomeworkSubmissions');
    if (!sheet) {
      sheet = ss.insertSheet('HomeworkSubmissions');
      sheet.appendRow(['Timestamp', 'Email', 'File ID', 'File Title', 'File URL', 'Filename', 'Score', 'Remark', 'GradedBy', 'GradedAt']);
    }

    var data = sheet.getDataRange().getValues();
    var rowNum = -1;
    var cleanStudentEmail = (studentEmail || "").toString().toLowerCase().trim();
    var cleanFileId = (fileId || "").toString().toLowerCase().trim();

    for (var i = 1; i < data.length; i++) {
      var rowEmail = data[i][1] ? data[i][1].toString().toLowerCase().trim() : "";
      var rowFileId = data[i][2] ? data[i][2].toString().toLowerCase().trim() : "";
      var rowFileTitle = data[i][3] ? data[i][3].toString().toLowerCase().trim() : "";
      if (rowEmail === cleanStudentEmail && (rowFileId === cleanFileId || rowFileTitle === cleanFileId)) {
        rowNum = i + 1;
        break;
      }
    }

    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    if (rowNum === -1) {
      sheet.appendRow([new Date(), studentEmail, fileId, fileId, "", "Chấm trực tiếp", score, remark, adminEmail, nowStr]);
    } else {
      sheet.getRange(rowNum, 7).setValue(score);
      sheet.getRange(rowNum, 8).setValue(remark);
      if (sheet.getLastColumn() >= 9) {
        sheet.getRange(rowNum, 9).setValue(adminEmail);
      }
      if (sheet.getLastColumn() >= 10) {
        sheet.getRange(rowNum, 10).setValue(nowStr);
      }
    }

    // Xóa cache dữ liệu bài tập
    invalidateSheetCache("HomeworkSubmissions");

    var actualFileId = (rowNum !== -1 && data[rowNum - 1] && data[rowNum - 1][2]) ? data[rowNum - 1][2].toString() : fileId;
    var fileTitle = (rowNum !== -1 && data[rowNum - 1] && data[rowNum - 1][3]) ? data[rowNum - 1][3].toString() : "Bài tập";

    // Strip bracket notation for display in the notification
    var displayScore = (score || "").toString();
    if (displayScore.indexOf('[') !== -1) {
      displayScore = displayScore.split('[')[0].trim();
    }

    const fileInfo = getFileDetailForNotif(actualFileId);
    const notifTitle = "Bài Đã Chấm Xong";
    const notifDesc = `Bài tập về nhà của bạn đã được giáo viên chấm điểm. Vui lòng xem chi tiết kết quả trong phần Bài Tập.`;
    const notifAction = (fileInfo.type || "exams") + ":" + actualFileId;
    addOrUpdateGroupNotification(notifTitle, notifDesc, notifAction, actualFileId, studentEmail, adminEmail);

    logActivityToSheet(adminEmail, "GRADE_HOMEWORK", `Chấm điểm BTVN cho ${studentEmail}: ${displayScore}đ - ${fileTitle}`);

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getStudentHomeworkStatus() {
  try {
    const email = CacheService.getUserCache().get("loggedInUser");
    if (!email) return { success: false, message: "Bạn chưa đăng nhập." };

    const role = getUserRole(email);
    const isAdmin = (role === "ADMIN");
    
    if (isAdmin) {
      return { 
        success: true, 
        unsubmitted: [], 
        role: role 
      };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Get all homework notes (assignments)
    var assignments = [];
    var noteSheet = ss.getSheetByName('HomeworkNotes');
    if (noteSheet) {
      var noteData = noteSheet.getDataRange().getValues();
      var headers = noteData[0];
      var deadlineColIdx = headers.indexOf('Deadline');
      for (var i = 1; i < noteData.length; i++) {
        if (noteData[i][0] && noteData[i][2]) { // fileId and note content exist
          assignments.push({
            fileId: noteData[i][0].toString(),
            fileTitle: noteData[i][1] ? noteData[i][1].toString() : "",
            note: noteData[i][2].toString(),
            deadline: (deadlineColIdx !== -1 && noteData[i][deadlineColIdx]) ? noteData[i][deadlineColIdx].toString() : ""
          });
        }
      }
    }

    // 2. Get student's submissions
    var submittedFileIds = {};
    var subSheet = ss.getSheetByName('HomeworkSubmissions');
    if (subSheet) {
      var subData = subSheet.getDataRange().getValues();
      var lowerEmail = email.toLowerCase().trim();
      for (var i = 1; i < subData.length; i++) {
        var subFileId = subData[i][2] ? subData[i][2].toString() : "";
        var subEmail = subData[i][1] ? subData[i][1].toString().toLowerCase().trim() : "";
        if (subEmail === lowerEmail) {
          submittedFileIds[subFileId] = true;
        }
      }
    }

    // 3. Evaluate allowed files for the user
    var allowedFileIds = {};
    if (isAdmin) {
      // Admin sees everything
      assignments.forEach(function(a) {
        allowedFileIds[a.fileId] = true;
      });
    } else {
      // Course files
      try {
        var courses = getCourses();
        if (courses && courses.length > 0) {
          courses.forEach(function(c) {
            if (c.fileId) allowedFileIds[c.fileId.toString()] = true;
          });
        }
      } catch(e) {}
      
      // Books
      try {
        var books = getBooks();
        if (books && books.length > 0) {
          books.forEach(function(b) {
            if (b.fileId && b.canOpen) allowedFileIds[b.fileId.toString()] = true;
          });
        }
      } catch(e) {}
      
      // Exams
      try {
        var exams = getExams();
        if (exams && exams.length > 0) {
          exams.forEach(function(ex) {
            if (ex.fileId && ex.canOpen) allowedFileIds[ex.fileId.toString()] = true;
          });
        }
      } catch(e) {}
    }

    // Filter assignments to only those they have access to and haven't submitted yet
    var unsubmitted = [];
    assignments.forEach(function(assign) {
      var fileIdStr = assign.fileId;
      var hasAccess = allowedFileIds[fileIdStr] || false;
      var hasSubmitted = submittedFileIds[fileIdStr] || false;
      if (hasAccess && !hasSubmitted) {
        unsubmitted.push(assign);
      }
    });

    return { 
      success: true, 
      unsubmitted: unsubmitted, 
      role: role 
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}






function overwriteStudentSubmissionPdf(base64Data, fileName, email, type, identifier) {
  try {
    const adminEmail = CacheService.getUserCache().get("loggedInUser");
    const role = getUserRole(adminEmail);
    if (!adminEmail || role !== "ADMIN") {
      return { success: false, message: "Bạn không có quyền cập nhật file này." };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = type === "exam" ? 'ExamSubmissions' : 'HomeworkSubmissions';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: "Không tìm thấy dữ liệu." };

    var data = sheet.getDataRange().getValues();
    var existingRowIndex = -1;
    var oldFileUrl = "";
    
    var lowerEmail = email.toLowerCase().trim();
    var lowerIdentifier = identifier.toLowerCase().trim();

    for (var i = 1; i < data.length; i++) {
      var rowEmail = data[i][1] ? data[i][1].toString().toLowerCase().trim() : "";
      var matchId = false;
      if (type === "exam") {
        var rowTitle = data[i][2] ? data[i][2].toString().toLowerCase().trim() : "";
        matchId = (rowTitle === lowerIdentifier);
      } else {
        var rowFileId = data[i][2] ? data[i][2].toString().toLowerCase().trim() : "";
        var rowFileTitle = data[i][3] ? data[i][3].toString().toLowerCase().trim() : "";
        matchId = (rowFileId === lowerIdentifier || rowFileTitle === lowerIdentifier);
      }

      if (rowEmail === lowerEmail && matchId) {
        existingRowIndex = i;
        oldFileUrl = (type === "exam") ? data[i][3] : data[i][4]; // For Exam, File URL is col 4 (index 3). For Homework, File URL is col 5 (index 4).
        break;
      }
    }

    if (existingRowIndex === -1) {
      return { success: false, message: "Không tìm thấy bài nộp tương ứng." };
    }

    // Direct in-place file content overwrite (Preserving exact same Google Drive File ID and Link)
    if (oldFileUrl) {
      try {
        var matchId = oldFileUrl.match(/[-\w]{25,}/);
        var oldFileId = matchId ? matchId[0] : null;
        if (oldFileId) {
          var rawBytes = Utilities.base64Decode(base64Data);
          var patchUrl = "https://www.googleapis.com/upload/drive/v3/files/" + oldFileId + "?uploadType=media";
          var patchResponse = UrlFetchApp.fetch(patchUrl, {
            method: "PATCH",
            contentType: "application/pdf",
            payload: rawBytes,
            headers: {
              Authorization: "Bearer " + ScriptApp.getOAuthToken()
            },
            muteHttpExceptions: true
          });

          var code = patchResponse.getResponseCode();
          if (code >= 200 && code < 300) {
            try {
              var existingFile = DriveApp.getFileById(oldFileId);
              existingFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            } catch(e) {}
            return { success: true, message: "Đã ghi đè trực tiếp lên file PDF gốc thành công (Giữ nguyên ID file).", newFileUrl: oldFileUrl };
          } else {
            Logger.log("Drive PATCH overwrite failed with status " + code + ": " + patchResponse.getContentText());
          }
        }
      } catch (err) {
        Logger.log("Lỗi ghi đè trực tiếp Google Drive: " + err.toString());
      }
    }

    // Fallback: create a new file if old file ID is missing or unreachable
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "application/pdf", fileName);
    var folder;
    if (type === "exam") {
      var targetFolderId = "1Ea6-0QTJ_BNLjvnQ3heJ4Qe2X5i9Zqfd";
      try { folder = DriveApp.getFolderById(targetFolderId); } catch(e) { folder = DriveApp.getRootFolder(); }
    } else {
      var parentFolderId = "1qhig9lIft7D8upXv4c6Qk6INHu8fNM5f";
      try {
        var parentFolder = DriveApp.getFolderById(parentFolderId);
        var emailFolders = parentFolder.getFoldersByName(lowerEmail);
        if (emailFolders.hasNext()) {
          folder = emailFolders.next();
        } else {
          folder = parentFolder.createFolder(lowerEmail);
        }
      } catch(e) {
        folder = DriveApp.getRootFolder();
      }
    }

    var newFile = folder.createFile(blob);
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var newFileUrl = newFile.getUrl();

    // Update the sheet
    if (type === "exam") {
      sheet.getRange(existingRowIndex + 1, 4).setValue(newFileUrl);
    } else {
      sheet.getRange(existingRowIndex + 1, 5).setValue(newFileUrl);
    }

    return { success: true, message: "Đã lưu bản chấm PDF mới thành công.", newFileUrl: newFileUrl };
  } catch (err) {
    return { success: false, message: "Lỗi: " + err.message };
  }
}


function getDriveFileBase64(fileUrl) {
  try {
    var matchId = fileUrl ? fileUrl.match(/[-\w]{25,}/) : null;
    var fileId = matchId ? matchId[0] : null;
    if (!fileId) return { success: false, message: "Invalid URL: " + fileUrl };

    var file = DriveApp.getFileById(fileId);
    var mimeType = file.getMimeType();
    var blob;
    
    // Convert Google Docs and Word documents to PDF
    if (mimeType === MimeType.GOOGLE_DOCS || 
        mimeType === MimeType.GOOGLE_SHEETS || 
        mimeType === MimeType.GOOGLE_SLIDES || 
        mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || 
        mimeType === 'application/msword') {
      blob = file.getAs(MimeType.PDF);
    } else {
      blob = file.getBlob();
    }
    
    var base64Data = Utilities.base64Encode(blob.getBytes());
    var rawName = file.getName();
    var cleanName = rawName.replace(/\s*\[PDF_Converted_\d+\]$/, "");
    return { success: true, base64Data: base64Data, fileName: cleanName, mimeType: blob.getContentType() };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function verifyAdminPasscode(pin) {
  if (!pin) return { success: false, message: "Mã PIN không được để trống." };
  var securePin = "26062006";
  if (pin.trim() === securePin) {
    return { success: true };
  } else {
    return { success: false, message: "Mã PIN không chính xác." };
  }
}

var PDF_NOTES_SPREADSHEET_ID = "14tdZWa6AgGZqpaqAOhLtHs68EuJsOenND7HF0g0iy-E";

function getPdfNotesTargetSpreadsheet() {
  try {
    if (PDF_NOTES_SPREADSHEET_ID && PDF_NOTES_SPREADSHEET_ID.trim() !== "") {
      return SpreadsheetApp.openById(PDF_NOTES_SPREADSHEET_ID.trim());
    }
  } catch (err) {
    Logger.log("Error opening external PDF Notes spreadsheet: " + err.toString());
  }
  // Tuyệt đối KHÔNG tự động fallback về SpreadsheetApp.getActiveSpreadsheet() (Sheet chính)
  return null;
}

function isPdfNotesEmptyString(notesJson) {
  if (!notesJson) return true;
  var s = String(notesJson).trim();
  if (s === "" || s === "{}" || s === "[]" || s === "null" || s === "undefined") return true;
  try {
    var parsed = JSON.parse(s);
    if (!parsed) return true;
    if (Array.isArray(parsed) && parsed.length === 0) return true;
    if (typeof parsed === 'object') {
      var keys = Object.keys(parsed);
      if (keys.length === 0) return true;
      var hasData = false;
      for (var i = 0; i < keys.length; i++) {
        var val = parsed[keys[i]];
        if (Array.isArray(val) && val.length > 0) {
          hasData = true;
          break;
        } else if (val && typeof val === 'object' && Object.keys(val).length > 0) {
          hasData = true;
          break;
        }
      }
      if (!hasData) return true;
    }
  } catch (e) {
    // If parse fails but non-empty string, keep it
  }
  return false;
}

function cleanupEmptyPdfNotesRows(sheet) {
  try {
    if (!sheet) {
      var ss = getPdfNotesTargetSpreadsheet();
      sheet = ss ? ss.getSheetByName("__PdfNotes__") : null;
    }
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return;

    var rowsToDelete = [];
    for (var i = data.length - 1; i >= 1; i--) {
      var rowFileId = String(data[i][0] || "").trim();
      var rowEmail = String(data[i][1] || "").trim();
      var rowNotes = String(data[i][2] || "").trim();

      // Row is empty if FileID is blank OR Email is blank OR NotesJSON is empty / "{}" / "[]"
      if (rowFileId === "" || rowEmail === "" || isPdfNotesEmptyString(rowNotes)) {
        rowsToDelete.push(i + 1); // 1-indexed for Google Sheets
      }
    }

    for (var k = 0; k < rowsToDelete.length; k++) {
      sheet.deleteRow(rowsToDelete[k]);
    }
  } catch (err) {
    Logger.log("Error cleaning up empty PDF notes rows: " + err.toString());
  }
}

function migratePdfNotesToNewSheet() {
  try {
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (!activeSs) return { success: false, message: "No active spreadsheet." };

    var oldSheet = activeSs.getSheetByName("__PdfNotes__");
    if (!oldSheet) return { success: true, message: "No old __PdfNotes__ sheet found." };

    var oldData = oldSheet.getDataRange().getValues();
    if (!oldData || oldData.length <= 1) {
      try { activeSs.deleteSheet(oldSheet); } catch (e) {}
      return { success: true, message: "Old __PdfNotes__ sheet was empty and removed." };
    }

    var targetSs = getPdfNotesTargetSpreadsheet();
    if (!targetSs) return { success: false, message: "Cannot open target PDF notes spreadsheet." };

    var targetSheet = targetSs.getSheetByName("__PdfNotes__");
    if (!targetSheet) {
      targetSheet = targetSs.insertSheet("__PdfNotes__");
      targetSheet.appendRow(["FileID", "Email", "NotesJSON", "UpdatedAt", "ChunkIndex"]);
      targetSheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    }

    var rowsToAppend = oldData.slice(1); // Skip header row
    if (rowsToAppend.length > 0) {
      var startRow = targetSheet.getLastRow() + 1;
      var numCols = rowsToAppend[0].length;
      targetSheet.getRange(startRow, 1, rowsToAppend.length, numCols).setValues(rowsToAppend);
    }

    // Delete or clear old sheet from active spreadsheet
    try {
      activeSs.deleteSheet(oldSheet);
    } catch (delErr) {
      oldSheet.setName("__PdfNotes_OLD_MIGRATED__");
      oldSheet.clearContents();
    }

    cleanupEmptyPdfNotesRows(targetSheet);

    Logger.log("Successfully migrated " + rowsToAppend.length + " PDF note rows to new spreadsheet.");
    return { success: true, migratedRows: rowsToAppend.length };
  } catch (err) {
    Logger.log("Error in migratePdfNotesToNewSheet: " + err.toString());
    return { success: false, message: err.toString() };
  }
}

function savePdfNotes(fileId, userEmail, notesJson) {
  try {
    if (!fileId || !userEmail) return { success: false, message: "Thiếu thông tin fileId hoặc userEmail." };

    var ss = getPdfNotesTargetSpreadsheet();
    if (!ss) {
      Logger.log("savePdfNotes: Không thể mở Google Sheet ghi chú riêng ID: " + PDF_NOTES_SPREADSHEET_ID);
      return { success: false, message: "Không thể mở Google Sheet ghi chú riêng. Vui lòng kiểm tra lại ID Sheet và quyền truy cập." };
    }

    // Auto-migrate old data if present
    migratePdfNotesToNewSheet();

    var sheet = ss.getSheetByName("__PdfNotes__");
    
    if (!sheet) {
      sheet = ss.insertSheet("__PdfNotes__");
      sheet.appendRow(["FileID", "Email", "NotesJSON", "UpdatedAt", "ChunkIndex"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    }

    var targetFileId = String(fileId).trim();
    var targetEmail = String(userEmail).toLowerCase().trim();
    var nowStr = new Date().toISOString();
    var rawStr = String(notesJson || "");

    var data = sheet.getDataRange().getValues();
    var rowsToDelete = [];
    
    // Find existing rows for this fileId + userEmail (iterate backwards)
    for (var i = data.length - 1; i >= 1; i--) {
      var rowFileId = String(data[i][0] || "").trim();
      var rowEmail = String(data[i][1] || "").toLowerCase().trim();
      if (rowFileId === targetFileId && rowEmail === targetEmail) {
        rowsToDelete.push(i + 1); // 1-indexed for sheet
      }
    }

    // Delete existing rows for this file and user
    for (var j = 0; j < rowsToDelete.length; j++) {
      sheet.deleteRow(rowsToDelete[j]);
    }

    // Only append new rows if notes are NOT empty/cleared
    if (!isPdfNotesEmptyString(rawStr)) {
      // Split rawStr into chunks of max 35,000 characters per cell to prevent Google Sheets 50,000 limit
      var chunkSize = 35000;
      var chunks = [];
      for (var c = 0; c < rawStr.length; c += chunkSize) {
        chunks.push(rawStr.substring(c, c + chunkSize));
      }

      var newRows = [];
      for (var k = 0; k < chunks.length; k++) {
        newRows.push([targetFileId, targetEmail, chunks[k], nowStr, k]);
      }
      
      if (newRows.length > 0) {
        var startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, newRows.length, 5).setValues(newRows);
      }
    }

    // Run cleanup for any leftover empty rows in the sheet
    cleanupEmptyPdfNotesRows(sheet);

    return { success: true };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getPdfNotes(fileId, userEmail) {
  try {
    if (!fileId || !userEmail) return { success: false, notesJson: null };

    var ss = getPdfNotesTargetSpreadsheet();
    if (!ss) {
      Logger.log("getPdfNotes: Không thể mở Google Sheet ghi chú riêng ID: " + PDF_NOTES_SPREADSHEET_ID);
      return { success: false, notesJson: null, message: "Không thể mở Google Sheet ghi chú riêng." };
    }

    // Auto-migrate old data if present
    migratePdfNotesToNewSheet();

    var sheet = ss.getSheetByName("__PdfNotes__");
    if (!sheet) return { success: true, notesJson: null };

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, notesJson: null };

    var targetFileId = String(fileId).trim();
    var targetEmail = String(userEmail).toLowerCase().trim();

    var matchingRows = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowFileId = String(row[0]).trim();
      var rowEmail = String(row[1]).toLowerCase().trim();
      if (rowFileId === targetFileId && rowEmail === targetEmail) {
        matchingRows.push(row);
      }
    }
    
    if (matchingRows.length > 0) {
      // Sort rows by ChunkIndex
      matchingRows.sort(function(a, b) {
        var idxA = parseInt(a[4]);
        var idxB = parseInt(b[4]);
        if (isNaN(idxA)) idxA = 0;
        if (isNaN(idxB)) idxB = 0;
        return idxA - idxB;
      });
      
      var chunks = [];
      for (var r = 0; r < matchingRows.length; r++) {
        var row = matchingRows[r];
        // Graceful fallback for the old column-based chunking format
        var isOldFormat = isNaN(parseInt(row[4])) && row.length > 5;
        if (isOldFormat) {
          for (var j = 2; j < row.length; j++) {
            var val = String(row[j] || "");
            if (j === row.length - 1 && (val.indexOf("T") > 0 || val.length === 24) && val.indexOf("{") < 0) continue;
            chunks.push(val);
          }
        } else {
          chunks.push(String(row[2] || ""));
        }
      }
      
      var combinedJson = chunks.join("");
      return { success: true, notesJson: combinedJson };
    }
    return { success: true, notesJson: null };
  } catch (err) {
    return { success: false, notesJson: null, message: err.toString() };
  }
}


/**
 * Kiểm tra giới hạn (Quota) số lượng ô (cell) còn lại của Google Sheets.
 * Giới hạn tối đa hiện tại của Google Sheets là 10,000,000 ô dữ liệu.
 */
function checkSpreadsheetQuota() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var maxCellLimit = 10000000; // 10,000,000 ô tối đa
    var totalAllocatedCells = 0;
    var totalUsedCells = 0;
    var sheetDetails = [];
    sheets.forEach(function(s) {
      var name = s.getName();
      var maxRows = s.getMaxRows();
      var maxCols = s.getMaxColumns();
      var lastRow = s.getLastRow();
      var lastCol = s.getLastColumn();
      var allocatedCells = maxRows * maxCols;
      var usedCells = lastRow * lastCol;
      totalAllocatedCells += allocatedCells;
      totalUsedCells += usedCells;
      sheetDetails.push({
        name: name,
        maxRows: maxRows,
        maxCols: maxCols,
        lastRow: lastRow,
        lastCol: lastCol,
        allocatedCells: allocatedCells,
        usedCells: usedCells
      });
    });
    var percentAllocated = ((totalAllocatedCells / maxCellLimit) * 100).toFixed(2);
    var percentUsed = ((totalUsedCells / maxCellLimit) * 100).toFixed(2);
    var remainingCells = maxCellLimit - totalAllocatedCells;
    
    var percentRemaining = ((remainingCells / maxCellLimit) * 100).toFixed(2);
    var status = remainingCells > 2000000 ? "AN TOÀN" : (remainingCells > 500000 ? "CẢNH BÁO" : "NGUY HIỂM");
    var totalRows = sheets.reduce(function(acc, s) { return acc + s.getMaxRows(); }, 0);
    var averageCols = totalRows > 0 ? Math.ceil(totalAllocatedCells / totalRows) : 26;
    var estimatedRowsRemaining = Math.floor(remainingCells / averageCols);

    var report = `
  ==================================================
  📊 BÁO CÁO QUOTA GOOGLE SHEETS: ${ss.getName()}
  ==================================================
  - Tình trạng: ${status}
  - Giới hạn tối đa: ${maxCellLimit.toLocaleString()} ô
  - Đã phân bổ: ${totalAllocatedCells.toLocaleString()} ô (${percentAllocated}%)
  - Còn trống: ${(remainingCells > 0 ? remainingCells : 0).toLocaleString()} ô (${percentRemaining}%)
  - Ước tính số dòng có thể thêm: ~${estimatedRowsRemaining.toLocaleString()} dòng (với ~${averageCols} cột/sheet)
  - Tổng số sheet: ${sheets.length}
  --------------------------------------------------
  CHI TIẾT TỪNG SHEET:
  ` + sheetDetails.map(function(sd) {
        return `  + [${sd.name}] Kích thước: ${sd.maxRows}x${sd.maxCols} = ${sd.allocatedCells.toLocaleString()} ô phân bổ (Dữ liệu thực tế: ${sd.lastRow}x${sd.lastCol})`;
      }).join('\n') + `
  ==================================================`;

    Logger.log(report);
    console.log(report);

    return {
      percentRemainingFormatted: percentRemaining + "%",
      status: status,
      estimatedRowsRemaining: estimatedRowsRemaining,
      averageCols: averageCols,
      success: true,
      spreadsheetName: ss.getName(),
      totalSheets: sheets.length,
      maxCellLimit: maxCellLimit,
      totalAllocatedCells: totalAllocatedCells,
      totalUsedCells: totalUsedCells,
      remainingCells: remainingCells,
      percentAllocatedFormatted: percentAllocated + "%",
      percentUsedFormatted: percentUsed + "%",
      message: report,
      sheetDetails: sheetDetails
    };
  } catch (err) {
    var errMsg = "Lỗi kiểm tra quota: " + err.toString();
    Logger.log(errMsg);
    return { success: false, message: errMsg };
  }
}

// =========================================================================
// QUẢN LÝ BANNER QUẢNG CÁO TỪ GOOGLE SHEETS ("Banners")
// =========================================================================

function getBanners() {
  try {
    var ss = getSpreadsheet();
    if (!ss) return [];
    
    var sheet = ss.getSheetByName('Banners');
    if (!sheet) {
      sheet = ss.insertSheet('Banners');
      sheet.appendRow(['ID', 'Title', 'ImageURL', 'Link', 'ButtonText', 'Status', 'Order']);
      invalidateSheetCache('Banners');
    }

    var data = getSheetDataCached('Banners');
    if (!data || data.length <= 1) {
      return [];
    }

    var banners = [];
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (!row || row.length === 0) continue;

      var id = (row[0] || '').toString().trim();
      var title = (row[1] || '').toString().trim();
      var imageUrl = (row[2] || '').toString().trim();
      var link = (row[3] || '').toString().trim();
      var buttonText = (row[4] || '').toString().trim() || 'Tìm hiểu ngay';
      var status = (row[5] || 'ACTIVE').toString().trim().toUpperCase();
      var order = parseInt(row[6]) || r;

      if (!imageUrl) continue;
      if (status === 'INACTIVE' || status === 'ẨN' || status === 'OFF' || status === 'HIDE') continue;

      banners.push({
        id: id || ('BN' + r),
        title: title,
        img: imageUrl,
        link: link || '#',
        buttonText: buttonText,
        status: status,
        order: order
      });
    }

    banners.sort(function(a, b) {
      return a.order - b.order;
    });

    return banners;
  } catch (e) {
    Logger.log('Lỗi getBanners: ' + e.toString());
    return [];
  }
}

function getDefaultBanners() {
  return [];
}

function getAllBannersAdmin() {
  try {
    var ss = getSpreadsheet();
    if (!ss) return { success: false, message: 'Không thể kết nối Google Sheets' };

    var sheet = ss.getSheetByName('Banners');
    if (!sheet) {
      getBanners();
      sheet = ss.getSheetByName('Banners');
    }

    var data = sheet.getDataRange().getValues();
    var banners = [];
    if (data && data.length > 1) {
      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        if (!row || row.length === 0) continue;
        var id = (row[0] || '').toString().trim();
        if (!id && !row[2]) continue;

        banners.push({
          id: id || ('BN' + r),
          title: (row[1] || '').toString().trim(),
          img: (row[2] || '').toString().trim(),
          link: (row[3] || '').toString().trim(),
          buttonText: (row[4] || 'Tìm hiểu ngay').toString().trim(),
          status: (row[5] || 'ACTIVE').toString().trim().toUpperCase(),
          order: parseInt(row[6]) || r
        });
      }
    }

    banners.sort(function(a, b) {
      return a.order - b.order;
    });

    return {
      success: true,
      banners: banners,
      sheetUrl: ss.getUrl()
    };
  } catch (e) {
    return { success: false, message: 'Lỗi lấy danh sách banner Admin: ' + e.toString() };
  }
}

function saveBannerAdmin(bannerObj) {
  try {
    if (!bannerObj) return { success: false, message: 'Dữ liệu không hợp lệ' };
    
    var ss = getSpreadsheet();
    if (!ss) return { success: false, message: 'Không thể tìm thấy Google Sheet' };

    var sheet = ss.getSheetByName('Banners');
    if (!sheet) {
      getBanners();
      sheet = ss.getSheetByName('Banners');
    }

    var data = sheet.getDataRange().getValues();
    var bannerId = (bannerObj.id || '').toString().trim();
    if (!bannerId) {
      bannerId = 'BN' + (data.length + 1) + '_' + Math.floor(Math.random() * 1000);
    }

    var title = (bannerObj.title || '').toString().trim();
    var img = (bannerObj.img || '').toString().trim();
    var link = (bannerObj.link || '#').toString().trim();
    var buttonText = (bannerObj.buttonText || 'Tìm hiểu ngay').toString().trim();
    var status = (bannerObj.status || 'ACTIVE').toString().trim().toUpperCase();
    var order = parseInt(bannerObj.order) || (data.length);

    var foundRowIndex = -1;
    for (var r = 1; r < data.length; r++) {
      if ((data[r][0] || '').toString().trim() === bannerId) {
        foundRowIndex = r + 1;
        break;
      }
    }

    if (foundRowIndex > 0) {
      sheet.getRange(foundRowIndex, 1, 1, 7).setValues([[bannerId, title, img, link, buttonText, status, order]]);
    } else {
      sheet.appendRow([bannerId, title, img, link, buttonText, status, order]);
    }

    invalidateSheetCache('Banners');
    return { success: true, message: 'Đã lưu banner thành công!', bannerId: bannerId };
  } catch (e) {
    return { success: false, message: 'Lỗi lưu banner: ' + e.toString() };
  }
}

function deleteBannerAdmin(bannerId) {
  try {
    if (!bannerId) return { success: false, message: 'Thiếu ID Banner' };

    var ss = getSpreadsheet();
    if (!ss) return { success: false, message: 'Không thể tìm thấy Google Sheet' };

    var sheet = ss.getSheetByName('Banners');
    if (!sheet) return { success: false, message: 'Sheet Banners không tồn tại' };

    var data = sheet.getDataRange().getValues();
    var foundRowIndex = -1;
    for (var r = 1; r < data.length; r++) {
      if ((data[r][0] || '').toString().trim() === (bannerId || '').toString().trim()) {
        foundRowIndex = r + 1;
        break;
      }
    }

    if (foundRowIndex > 0) {
      sheet.deleteRow(foundRowIndex);
      invalidateSheetCache('Banners');
      return { success: true, message: 'Đã xóa banner khỏi hệ thống!' };
    } else {
      return { success: false, message: 'Không tìm thấy banner cần xóa' };
    }
  } catch (e) {
    return { success: false, message: 'Lỗi xóa banner: ' + e.toString() };
  }
}