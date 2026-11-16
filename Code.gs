
const CONFIG = {
  CAPACITY: 999,
  SHEET_NAME: "Registrations",
  SPREADSHEET_PROPERTY: "REASONIQ99_SPREADSHEET_ID",
  CATEGORIES: ["UPSC & State PSC", "SSC", "Banking", "Railways", "Defence", "Other Competitive Exams"]
};

function setupReasonIQ99() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(CONFIG.SPREADSHEET_PROPERTY);
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create("ReasonIQ99 Founding Batch Registrations");
  if (!id) props.setProperty(CONFIG.SPREADSHEET_PROPERTY, ss.getId());
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(["Submitted At","Reference ID","Full Name","Mobile Number","Email","State","Exam Category","Consent"]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,8).setFontWeight("bold");
    sh.autoResizeColumns(1,8);
  }
  Logger.log("Your private registration sheet: " + ss.getUrl());
  return ss.getUrl();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("ReasonIQ99 Founding Batch Registration")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getRegistrationStats() {
  const sh = getSheet_();
  const registered = Math.max(0, sh.getLastRow() - 1);
  return {registered, capacity: CONFIG.CAPACITY, remaining: Math.max(0, CONFIG.CAPACITY-registered), isFull: registered >= CONFIG.CAPACITY};
}

function submitRegistration(f) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return {success:false, message:"System busy. Please try again."};
  try {
    f = f || {};
    if (String(f.website || "").trim()) return {success:false, message:"Unable to process submission."};
    const name = clean_(f.fullName,100);
    const mobile = mobile_(f.mobile);
    const email = String(f.email || "").trim().toLowerCase();
    const state = clean_(f.state,80);
    const category = clean_(f.category,80);
    if (name.length < 2) return fail_("Enter your full name.");
    if (!/^[6-9]\d{9}$/.test(mobile)) return fail_("Enter a valid 10-digit Indian mobile number.");
    if (email && !/^[^\s@]+@[^\s@]+\\.[^\s@]+$/.test(email)) return fail_("Enter a valid email address.");
    if (!state) return fail_("Enter your state.");
    if (!CONFIG.CATEGORIES.includes(category)) return fail_("Select a valid exam category.");
    if (String(f.consent) !== "YES") return fail_("Please accept the Terms & Conditions and Privacy Policy.");

    const sh = getSheet_(), last = sh.getLastRow(), count = Math.max(0,last-1);
    if (count >= CONFIG.CAPACITY) return {success:false,full:true,message:"All 999 seats are filled. Registration is closed."};

    if (last >= 2) {
      const old = sh.getRange(2,4,last-1,2).getDisplayValues();
      for (const row of old) {
        if (mobile_(row[0]) === mobile) return {success:false,duplicate:true,message:"This mobile number is already registered."};
        if (email && String(row[1] || "").trim().toLowerCase() === email) return {success:false,duplicate:true,message:"This email address is already registered."};
      }
    }
    const n = count + 1;
    const ref = "RI99-" + Utilities.formatString("%06d",n);
    sh.appendRow([new Date(),ref,name,mobile,email,state,category,"YES"]);
    return {success:true,reference:ref,registered:n,remaining:CONFIG.CAPACITY-n,message:"Registration successful!"};
  } catch(e) {
    console.error(e);
    return {success:false,message:"Something went wrong. Please try again."};
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(CONFIG.SPREADSHEET_PROPERTY);
  if (!id) throw new Error("Run setupReasonIQ99() first.");
  const sh = SpreadsheetApp.openById(id).getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error("Registration sheet missing. Run setupReasonIQ99() again.");
  return sh;
}
function mobile_(v) {
  let s = String(v || "").replace(/\D/g,"");
  if (s.length===12 && s.startsWith("91")) s=s.slice(2);
  if (s.length===11 && s.startsWith("0")) s=s.slice(1);
  return s;
}
function clean_(v,n) { return String(v || "").replace(/[<>]/g,"").trim().slice(0,n); }
function fail_(message) { return {success:false,message}; }
