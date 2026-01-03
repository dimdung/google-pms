/*********************************************************
 * Lee High Inn – Enhanced Script (Header-Driven)
 * Features:
 * 1) F is now "Number of Night(s)" and drives Subtotal/Total
 * 2) Subtotal/Total auto-update when Nights/Amount/Tax Rate changes
 * 3) CheckIn/CheckOut timestamps reliably stamp again
 * 4) Repairs missing headers (U/V blank titles issue)
 * 5) Adds menu: Lee High Inn → Generate/Reprint Invoice
 * 6) Optionally hides Duration and TempNote (instead of deleting)
 * 7) Housekeeping automation (HK Done timestamp, color coding)
 * 8) Payment Type handling and validation
 * 9) Daily/weekly reporting functions
 * 10) Room availability dashboard
 * 11) Guest history lookup
 * 12) Bulk operations menu items
 * 13) Enhanced error handling and logging
 * 14) Data validation improvements
 *********************************************************/

const CFG = {
  SHEET: "FrontDesk_Log",
  TIMEZONE: "America/New_York",

  MOTEL: {
    name: "Lee High Inn Motel",
    addr1: "9865 Fairfax Blvd",
    addr2: "Fairfax, VA 22030",
    phone: "703-975-5067",
    email: "saatgdllc@gmail.com",
  },

  DEFAULT_TAX_RATE: 0.13,

  COLOR: {
    CHECKIN_YES: "#C6EFCE",        // Light green
    CHECKOUT_YES: "#FFE699",       // Light yellow/amber (better visibility)
    HK_READY: "#FFF2CC",            // Light yellow
    HK_DONE: "#D5E8D4",             // Light green
    PAYMENT_CASH: "#E1D5E7",        // Light purple
    PAYMENT_CARD: "#D5E8D4",       // Light green
    PAYMENT_OTHER: "#FFF2CC",       // Light yellow
    ROOM_OCCUPIED: "#FFC7CE",       // Light red
    ROOM_AVAILABLE: "#C6EFCE",      // Light green
    ROOM_MAINTENANCE: "#D9D9D9",   // Light gray
    ROOM_READY: "#FFF2CC",          // Light yellow
    OLD_DATA: "#E8E8E8",            // Light gray for old/inactive data
    OLD_DATA_TEXT: "#999999",       // Gray text for old/inactive data
  },

  HK_READY_TEXT: "Ready for Cleaning",
  HK_DONE_TEXT: "Cleaned - ReadyFor Rent",
  
  PAYMENT_TYPES: ["Cash", "Credit Card", "Debit Card", "Check", "Other"],
  
  // Column constants
  OLD_DATA_END_COLUMN: 18, // Column R - last column to gray out for old data
  
  LOG_ENABLED: true,
};

/* ===================== MENU ===================== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Lee High Inn");
  
  menu.addItem("Generate/Reprint Invoice (Selected Row)", "menuGenerateInvoice");
  menu.addSeparator();
  
  menu.addSubMenu(ui.createMenu("Reports")
    .addItem("Daily Report", "menuDailyReport")
    .addItem("Weekly Report", "menuWeeklyReport")
    .addItem("Monthly Report", "menuMonthlyReport")
    .addItem("Room Occupancy Report", "menuRoomOccupancyReport"));
  
  menu.addSubMenu(ui.createMenu("Dashboard")
    .addItem("Unified Room Dashboard (60 Rooms)", "menuUnifiedRoomDashboard")
    .addItem("Update Daily Dashboard", "menuUpdateDailyDashboard")
    .addSeparator()
    .addItem("Room Availability", "menuRoomAvailability")
    .addItem("Today's Check-ins", "menuTodaysCheckins")
    .addItem("Today's Check-outs", "menuTodaysCheckouts")
    .addItem("Pending Housekeeping", "menuPendingHousekeeping"));
  
  menu.addSubMenu(ui.createMenu("Guest")
    .addItem("Lookup Guest History", "menuGuestHistory"));
  
  menu.addSubMenu(ui.createMenu("Bulk Operations")
    .addItem("Generate Invoices for Selected Rows", "menuBulkGenerateInvoices")
    .addItem("Update Tax Rate for Selected Rows", "menuBulkUpdateTaxRate")
    .addItem("Mark Selected as Checked Out", "menuBulkCheckout")
    .addSeparator()
    .addItem("Refresh Old Row Styling", "menuRefreshOldRowStyling"));
  
  menu.addSeparator();
  menu.addItem("Setup/Repair Headers", "menuSetupOnce");
  menu.addItem("Manage Room Maintenance", "menuManageRoomMaintenance");
  
  menu.addToUi();
}

/* ===================== RUN ONCE ===================== */
function setupOnce() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
  if (!sh) throw new Error(`Sheet not found: ${CFG.SHEET}`);

  repairHeaders_(sh);
  createInstallableOnEditTrigger_();
  protectTimestampColumns_(sh);

  // Optional: hide columns you don't want staff using
  hideIfExists_(sh, "Duration");
  hideIfExists_(sh, "TempNote");

  SpreadsheetApp.getUi().alert("Setup complete. Reload the sheet.");
}

/* ===================== TRIGGERS ===================== */
function createInstallableOnEditTrigger_() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "handleEdit");
  if (exists) return;

  ScriptApp.newTrigger("handleEdit")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

// keep simple trigger delegating
function onEdit(e) { handleEdit(e); }

/* ===================== EDIT HANDLER ===================== */
function handleEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== CFG.SHEET) return;

    const row = e.range.getRow();
    if (row < 2) return;

    // Always keep headers healthy (handles your "U/V blank title" issue)
    if (row === 1) repairHeaders_(sh);

    const col = e.range.getColumn();
    const v = (e.value || "").toString().trim();
    const vBool = toYes_(v);

    // Resolve columns by header (robust even if columns shift)
    const C = cols_(sh);

    // If Total With Tax is entered directly, calculate backwards (flat amount mode)
    if (col === C.total && v && !isNaN(toNumber_(v))) {
      updateQuoteFromTotal_(sh, row, C, toNumber_(v));
      return;
    }

    // Update quote totals when Amount / Nights / Tax changes
    if ([C.amount, C.nights, C.taxRate].includes(col)) {
      updateQuoteForRow_(sh, row, C);
    }

    // Payment Type handling
    if (col === C.paymentType && v) {
      handlePaymentTypeChange_(sh, row, C, v);
    }

    // CheckIn timestamp
    if (col === C.checkIn && vBool) {
      if (!sh.getRange(row, C.checkInTime).getValue()) {
        sh.getRange(row, C.checkInTime).setValue(nowEST_());
      }
      sh.getRange(row, C.checkIn).setBackground(CFG.COLOR.CHECKIN_YES);
      updateQuoteForRow_(sh, row, C);
      
      // Clear "Cleaned - ReadyFor Rent" status from older rows for the same room
      // since the room is now occupied by a new guest
      const room = sh.getRange(row, C.room).getValue();
      if (room && C.hkStatus) {
        clearOldCleanedStatus_(sh, room, row, C);
      }
      
      // Gray out Date and CheckOut columns for old rows of the same room
      if (room) {
        grayOutOldRows_(sh, room, row, C);
      }
      
      log_("Check-in processed", { room: sh.getRange(row, C.room).getValue(), row });
      return;
    }

  // CheckOut timestamp + HK status + invoice
  if (col === C.checkOut && vBool) {
    if (!sh.getRange(row, C.checkOutTime).getValue()) {
      sh.getRange(row, C.checkOutTime).setValue(nowEST_());
    }
    sh.getRange(row, C.checkOut).setBackground(CFG.COLOR.CHECKOUT_YES);

    // Ready for housekeeping
    if (C.hkStatus) {
      sh.getRange(row, C.hkStatus).setValue(CFG.HK_READY_TEXT).setBackground(CFG.COLOR.HK_READY);
    }

    updateQuoteForRow_(sh, row, C);
    generateInvoiceForRow_(sh, row, C, false);
    
    // Protect Total With Tax column (H column) after checkout
    protectTotalAfterCheckout_(sh, row, C);
    
    // Update Daily Dashboard (non-blocking, runs in background)
    try {
      updateDailyDashboard_();
    } catch (error) {
      log_("Error updating dashboard on checkout", { error: error.toString() });
      // Don't block checkout if dashboard update fails
    }
    
    log_("Check-out processed", { room: sh.getRange(row, C.room).getValue(), row });
    return;
  }
  
  // Prevent editing Total With Tax if already checked out
  if (col === C.total && C.checkOut) {
    const isCheckedOut = toYes_(sh.getRange(row, C.checkOut).getValue());
    if (isCheckedOut) {
      SpreadsheetApp.getUi().alert("Cannot modify Total With Tax after checkout. Please void the invoice first if correction is needed.");
      // Revert the change
      const oldValue = sh.getRange(row, C.total).getValue();
      SpreadsheetApp.flush();
      sh.getRange(row, C.total).setValue(oldValue);
      return;
    }
  }
  
  // Prevent editing timestamp columns
  if ([C.checkInTime, C.checkOutTime, C.cleanedTime].includes(col)) {
    SpreadsheetApp.getUi().alert("This column is protected and cannot be edited. Timestamps are set automatically.");
    // Revert the change
    const oldValue = sh.getRange(row, col).getValue();
    SpreadsheetApp.flush();
    sh.getRange(row, col).setValue(oldValue);
    return;
  }

    // HK Done timestamp and status update
    if (col === C.hkDone && vBool) {
      handleHKDone_(sh, row, C);
      return;
    }

  } catch (error) {
    log_("Error in handleEdit", { error: error.toString(), stack: error.stack });
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

/* ===================== COLUMN RESOLUTION ===================== */
function cols_(sh) {
  // Required headers (we will auto-create/repair these)
  const map = {
    date: "Date",
    room: "Room #",
    guest: "Full Name",
    amount: "Amount",
    nights: "Number of Night(s)",
    subtotal: "Subtotal",
    total: "Total With Tax",
    paymentType: "Payment Type",
    checkIn: "CheckIn",
    checkInTime: "CheckInTime",
    checkOut: "CheckOut",
    checkOutTime: "CheckOutTime",
    hkStatus: "HK Status",
    hkDone: "HK Done",
    cleanedTime: "CleanedTime",
    deskNotes: "Desk Notes",
    hkNotes: "HK Notes",
    taxRate: "Tax Rate",
    guestEmail: "Guest Email",
    processor: "Payment Processor",
    receipt: "Processor Receipt #",
    last4: "Card Last4",
    auth: "Auth Code",
    invoiceNo: "Invoice #",
    invoiceStatus: "Invoice Status",
    invoiceUrl: "Invoice PDF URL",
  };

  const headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => (h || "").toString().trim());

  const idx = {};
  for (const [k, name] of Object.entries(map)) {
    // Try exact match first
    let i = headerRow.indexOf(name);
    
    // If not found, try case-insensitive match
    if (i < 0) {
      i = headerRow.findIndex(h => h.toLowerCase() === name.toLowerCase());
    }
    
    // Special handling for CheckOut/Checkout variations
    if (i < 0 && (k === "checkOut" || k === "checkIn")) {
      const variations = k === "checkOut" 
        ? ["Checkout", "checkout", "Check Out", "check out"]
        : ["Checkin", "checkin", "Check In", "check in"];
      for (const variant of variations) {
        i = headerRow.findIndex(h => h.toLowerCase() === variant.toLowerCase());
        if (i >= 0) break;
      }
    }
    
    if (i >= 0) idx[k] = i + 1;
    else idx[k] = null; // may be optional like HK Status
  }
  return idx;
}

/* ===================== HEADER REPAIR ===================== */
function repairHeaders_(sh) {
  // Fix common broken state:
  // - Quoted Nights should be renamed to Number of Night(s)
  // - Some columns may be blank (U/V "no title" issue)
  // - Ensure required headers exist; do not move columns here

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  // Rename "Quoted Nights" -> "Number of Night(s)"
  for (let c = 1; c <= lastCol; c++) {
    const h = (headers[c-1] || "").toString().trim();
    if (h === "Quoted Nights") sh.getRange(1, c).setValue("Number of Night(s)");
  }

  // If there are blank headers, try to set them based on nearby known structure
  // At minimum, ensure Tax Rate / Guest Email / Payment Processor / Receipt exist.
  const required = ["Tax Rate", "Guest Email", "Payment Processor", "Processor Receipt #", "Invoice #", "Invoice Status", "Invoice PDF URL", "Payment Type", "HK Done", "CleanedTime"];
  const headerStrings = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => (x || "").toString().trim());

  // Helper function to add missing headers
  let refreshed = headerStrings.slice();
  const missingHeaders = [];
  
  if (!refreshed.includes("Tax Rate")) missingHeaders.push("Tax Rate");
  if (!refreshed.includes("Guest Email")) missingHeaders.push("Guest Email");
  if (!refreshed.includes("Payment Type")) missingHeaders.push("Payment Type");
  if (!refreshed.includes("HK Done")) missingHeaders.push("HK Done");
  if (!refreshed.includes("CleanedTime")) missingHeaders.push("CleanedTime");
  
  for (const header of missingHeaders) {
    const blank = refreshed.findIndex(x => x === "");
    if (blank >= 0) {
      sh.getRange(1, blank + 1).setValue(header);
      refreshed[blank] = header;
    }
  }

  // Also fix missing "Invoice Status" header if your sheet used "VOID" column instead
  const refreshed2 = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => (x || "").toString().trim());
  if (refreshed2.includes("VOID") && !refreshed2.includes("Invoice Status")) {
    // Rename VOID -> Invoice Status (matches the script)
    const c = refreshed2.indexOf("VOID") + 1;
    sh.getRange(1, c).setValue("Invoice Status");
  }

  // If your guest header is "Name" instead of "Full Name", rename it
  const refreshed3 = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => (x || "").toString().trim());
  if (refreshed3.includes("Name") && !refreshed3.includes("Full Name")) {
    const c = refreshed3.indexOf("Name") + 1;
    sh.getRange(1, c).setValue("Full Name");
  }
}

/* ===================== PROTECTION ===================== */
function protectTimestampColumns_(sh) {
  const headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => (x || "").toString().trim());

  const inIdx = headerRow.indexOf("CheckInTime");
  const outIdx = headerRow.indexOf("CheckOutTime");
  const cleanedIdx = headerRow.indexOf("CleanedTime");
  
  const colsToProtect = [];
  if (inIdx >= 0) colsToProtect.push(inIdx + 1);
  if (outIdx >= 0) colsToProtect.push(outIdx + 1);
  if (cleanedIdx >= 0) colsToProtect.push(cleanedIdx + 1);
  
  if (colsToProtect.length === 0) return;

  const protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  colsToProtect.forEach(c => {
    const rng = sh.getRange(1, c, sh.getMaxRows(), 1);
    const a1 = rng.getA1Notation();
    if (protections.some(p => p.getRange().getA1Notation() === a1)) return;

    const p = rng.protect();
    p.setDescription("Auto timestamps (do not edit)");
    p.setWarningOnly(false);

    const editors = p.getEditors();
    if (editors && editors.length) p.removeEditors(editors);
    if (p.canDomainEdit()) p.setDomainEdit(false);
  });
}

function protectTotalAfterCheckout_(sh, row, C) {
  // Protect Total With Tax column for this row after checkout
  if (!C.total) return;
  
  try {
    const protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    const totalCol = C.total;
    const rng = sh.getRange(row, totalCol);
    const a1 = rng.getA1Notation();
    
    // Check if already protected
    if (protections.some(p => p.getRange().getA1Notation() === a1)) return;
    
    const p = rng.protect();
    p.setDescription(`Total locked after checkout (Row ${row})`);
    p.setWarningOnly(false);
    
    const editors = p.getEditors();
    if (editors && editors.length) p.removeEditors(editors);
    if (p.canDomainEdit()) p.setDomainEdit(false);
    
    log_("Protected Total column after checkout", { row, column: totalCol });
  } catch (error) {
    log_("Error protecting Total column", { error: error.toString(), row });
  }
}

function hideIfExists_(sh, headerName) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => (x || "").toString().trim());
  const idx = headers.indexOf(headerName);
  if (idx >= 0) sh.hideColumns(idx + 1);
}

/* ===================== QUOTE CALC ===================== */
function updateQuoteForRow_(sh, row, C) {
  if (!C.room || !C.amount || !C.nights || !C.subtotal || !C.total || !C.taxRate) return;

  const room = sh.getRange(row, C.room).getValue();
  if (!room) return;

  const nights = Math.max(1, Math.floor(toNumber_(sh.getRange(row, C.nights).getValue() || 1)));
  const rate = toNumber_(sh.getRange(row, C.amount).getValue());
  const taxRate = toTaxRate_(sh.getRange(row, C.taxRate).getValue());

  const subtotal = +(rate * nights).toFixed(2);
  const total = +(subtotal * (1 + taxRate)).toFixed(2);

  sh.getRange(row, C.nights).setValue(nights);      // normalize
  sh.getRange(row, C.subtotal).setValue(subtotal);
  sh.getRange(row, C.total).setValue(total);
}

/* ===================== FLAT AMOUNT CALC (Total Including Tax) ===================== */
function updateQuoteFromTotal_(sh, row, C, flatTotal) {
  // User entered a flat amount (total including tax) - calculate backwards
  if (!C.room || !C.amount || !C.nights || !C.subtotal || !C.total || !C.taxRate) return;

  const room = sh.getRange(row, C.room).getValue();
  if (!room) return;

  const nights = Math.max(1, Math.floor(toNumber_(sh.getRange(row, C.nights).getValue() || 1)));
  const taxRate = toTaxRate_(sh.getRange(row, C.taxRate).getValue());

  // Calculate backwards from total
  // total = subtotal * (1 + taxRate)
  // subtotal = total / (1 + taxRate)
  const subtotal = +(flatTotal / (1 + taxRate)).toFixed(2);
  
  // Calculate per-night rate
  const rate = +(subtotal / nights).toFixed(2);

  // Update all values
  sh.getRange(row, C.nights).setValue(nights);
  sh.getRange(row, C.amount).setValue(rate);
  sh.getRange(row, C.subtotal).setValue(subtotal);
  sh.getRange(row, C.total).setValue(flatTotal);
  
  log_("Flat amount calculated", { 
    room, 
    flatTotal, 
    subtotal, 
    rate, 
    nights, 
    taxRate: (taxRate * 100).toFixed(2) + "%" 
  });
}

/* ===================== INVOICE ===================== */
function menuGenerateInvoice() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== CFG.SHEET) {
    SpreadsheetApp.getUi().alert("Run this from FrontDesk_Log.");
    return;
  }
  const row = sh.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert("Select a data row first.");
    return;
  }
  const C = cols_(sh);
  updateQuoteForRow_(sh, row, C);
  generateInvoiceForRow_(sh, row, C, true);
  SpreadsheetApp.getUi().alert("Invoice generated/reprinted.");
}

function generateInvoiceForRow_(sh, row, C, force) {
  if (!C.room || !C.guest || !C.amount || !C.nights || !C.taxRate || !C.invoiceNo || !C.invoiceStatus || !C.invoiceUrl || !C.checkOutTime) return;

  const room = sh.getRange(row, C.room).getValue();
  const guest = sh.getRange(row, C.guest).getValue();
  if (!room) return;

  let inv = sh.getRange(row, C.invoiceNo).getValue();
  if (!inv) {
    inv = nextInvoiceNo_();
    sh.getRange(row, C.invoiceNo).setValue(inv);
  }

  let status = sh.getRange(row, C.invoiceStatus).getValue();
  if (!status) {
    status = "PAID";
    sh.getRange(row, C.invoiceStatus).setValue(status);
  }
  if (String(status).toUpperCase() === "VOID") return;

  const existingUrl = sh.getRange(row, C.invoiceUrl).getValue();
  if (existingUrl && !force) return;

  const nights = Math.max(1, Math.floor(toNumber_(sh.getRange(row, C.nights).getValue() || 1)));
  const rate = toNumber_(sh.getRange(row, C.amount).getValue());
  const taxRate = toTaxRate_(sh.getRange(row, C.taxRate).getValue());

  const subtotal = +(rate * nights).toFixed(2);
  const tax = +(subtotal * taxRate).toFixed(2);
  const total = +(subtotal + tax).toFixed(2);

  const paidAt = sh.getRange(row, C.checkOutTime).getValue() || nowEST_();

  const processor = C.processor ? sh.getRange(row, C.processor).getValue() : "";
  const receipt = C.receipt ? sh.getRange(row, C.receipt).getValue() : "";
  const last4 = C.last4 ? sh.getRange(row, C.last4).getValue() : "";
  const auth = C.auth ? sh.getRange(row, C.auth).getValue() : "";
  const paymentType = C.paymentType ? sh.getRange(row, C.paymentType).getValue() : "";

  const perNightRate = rate.toFixed(2);
  const nightsText = nights === 1 ? "night" : "nights";
  
  const html = `
  <html><body style="font-family:Arial, sans-serif; font-size:13px; margin:0; padding:20px; background-color:#ffffff;">
    <div style="text-align:center; margin-bottom:25px; border-bottom:2px solid #333; padding-bottom:15px;">
      <h1 style="margin:0; color:#2c3e50; font-size:24px;">${CFG.MOTEL.name}</h1>
      <div style="margin-top:8px; color:#555; font-size:13px;">
        ${CFG.MOTEL.addr1}<br>
        ${CFG.MOTEL.addr2}
      </div>
      <div style="margin-top:8px; color:#666; font-size:12px;">
        Phone: ${CFG.MOTEL.phone} | Email: ${CFG.MOTEL.email}
      </div>
    </div>
    
    <div style="margin-bottom:20px;">
      <h2 style="margin:0 0 15px 0; color:#2c3e50; font-size:18px; border-bottom:1px solid #ddd; padding-bottom:8px;">INVOICE</h2>
      <table style="width:100%; margin-bottom:15px; border-collapse:collapse;">
        <tr>
          <td style="padding:5px 10px 5px 0; color:#555; width:140px;"><b>Invoice #:</b></td>
          <td style="padding:5px 0; color:#333;">${inv}</td>
        </tr>
        <tr>
          <td style="padding:5px 10px 5px 0; color:#555;"><b>Guest:</b></td>
          <td style="padding:5px 0; color:#333;">${guest || "-"}</td>
        </tr>
        <tr>
          <td style="padding:5px 10px 5px 0; color:#555;"><b>Room:</b></td>
          <td style="padding:5px 0; color:#333;">${room}</td>
        </tr>
        <tr>
          <td style="padding:5px 10px 5px 0; color:#555;"><b>Date (EST):</b></td>
          <td style="padding:5px 0; color:#333;">${fmt_(paidAt,"MM/dd/yyyy hh:mm a")}</td>
        </tr>
      </table>
    </div>

    <div style="margin-bottom:20px; padding:12px; background-color:#f8f9fa; border-left:4px solid #3498db;">
      <div style="font-size:12px; color:#555; margin-bottom:8px;"><b>Payment Information:</b></div>
      <table style="width:100%; font-size:12px; border-collapse:collapse;">
        <tr>
          <td style="padding:3px 10px 3px 0; color:#666; width:160px;">Payment Type:</td>
          <td style="padding:3px 0; color:#333;">${paymentType || "-"}</td>
        </tr>
        <tr>
          <td style="padding:3px 10px 3px 0; color:#666;">Payment Processor:</td>
          <td style="padding:3px 0; color:#333;">${processor || "-"}</td>
        </tr>
        <tr>
          <td style="padding:3px 10px 3px 0; color:#666;">Receipt / Transaction #:</td>
          <td style="padding:3px 0; color:#333;">${receipt || "-"}</td>
        </tr>
        <tr>
          <td style="padding:3px 10px 3px 0; color:#666;">Card Last 4:</td>
          <td style="padding:3px 0; color:#333;">${last4 || "-"}</td>
        </tr>
        <tr>
          <td style="padding:3px 10px 3px 0; color:#666;">Auth Code:</td>
          <td style="padding:3px 0; color:#333;">${auth || "-"}</td>
        </tr>
      </table>
    </div>

    <div style="margin-bottom:20px;">
      <table border="1" width="100%" cellpadding="10" cellspacing="0" style="border-collapse:collapse; border:1px solid #ddd;">
        <thead>
          <tr style="background-color:#34495e; color:#fff;">
            <th align="left" style="padding:10px; font-weight:bold; font-size:13px;">Description</th>
            <th align="right" style="padding:10px; font-weight:bold; font-size:13px; width:120px;">Amount</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="padding:10px; border-bottom:1px solid #eee;">
              <div style="font-weight:500; color:#2c3e50; margin-bottom:4px;">Lodging</div>
              <div style="font-size:11px; color:#666;">
                ${nights} ${nightsText} @ $${perNightRate}/night
              </div>
            </td>
            <td align="right" style="padding:10px; border-bottom:1px solid #eee; font-weight:500; color:#2c3e50;">
              $${subtotal.toFixed(2)}
            </td>
          </tr>
          <tr>
            <td style="padding:10px; border-bottom:1px solid #eee;">
              <div style="font-weight:500; color:#2c3e50;">Tax</div>
              <div style="font-size:11px; color:#666;">
                ${(taxRate*100).toFixed(2)}%
              </div>
            </td>
            <td align="right" style="padding:10px; border-bottom:1px solid #eee; font-weight:500; color:#2c3e50;">
              $${tax.toFixed(2)}
            </td>
          </tr>
          <tr style="background-color:#ecf0f1; font-weight:bold;">
            <td style="padding:12px; font-size:14px; color:#2c3e50;">Total Paid</td>
            <td align="right" style="padding:12px; font-size:16px; color:#27ae60;">
              $${total.toFixed(2)}
            </td>
          </tr>
        </tbody>
      </table>
    </div>

    <div style="margin-top:25px; padding-top:15px; border-top:1px solid #ddd; text-align:center;">
      <p style="font-size:11px; color:#7f8c8d; margin:0;">
        Receipt for lodging paid. Keep this document for your records.
      </p>
    </div>
  </body></html>`;

  const fileName = `${room}_${safeName_(guest)}_${fmt_(paidAt,"yyyyMMdd_HHmmss")}.pdf`;
  const pdf = Utilities.newBlob(html, "text/html").getAs("application/pdf").setName(fileName);
  const file = folderToday_().createFile(pdf);

  sh.getRange(row, C.invoiceUrl).setValue(file.getUrl());

  // Email optional
  if (C.guestEmail) {
    const email = sh.getRange(row, C.guestEmail).getValue();
    if (email) {
      GmailApp.sendEmail(email, `Invoice ${inv} - ${CFG.MOTEL.name}`, "Invoice attached.", { attachments: [pdf] });
    }
  }
}

/* ===================== DRIVE + SEQ ===================== */
function folderToday_() {
  // Create folder structure: PMS/Invoices/YYYY-MM-DD
  const root = DriveApp.getRootFolder();
  
  // Get or create PMS folder
  let pmsFolder;
  const pmsFolders = root.getFoldersByName("PMS");
  if (pmsFolders.hasNext()) {
    pmsFolder = pmsFolders.next();
  } else {
    pmsFolder = root.createFolder("PMS");
  }
  
  // Get or create Invoices folder inside PMS
  let invoicesFolder;
  const invoicesFolders = pmsFolder.getFoldersByName("Invoices");
  if (invoicesFolders.hasNext()) {
    invoicesFolder = invoicesFolders.next();
  } else {
    invoicesFolder = pmsFolder.createFolder("Invoices");
  }
  
  // Get or create date folder inside Invoices
  const dateName = fmt_(nowEST_(), "yyyy-MM-dd");
  const dateFolders = invoicesFolder.getFoldersByName(dateName);
  if (dateFolders.hasNext()) {
    return dateFolders.next();
  } else {
    return invoicesFolder.createFolder(dateName);
  }
}

function nextInvoiceNo_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const props = PropertiesService.getScriptProperties();
    const current = parseInt(props.getProperty("INVOICE_SEQ") || "0", 10);
    const next = current + 1;
    props.setProperty("INVOICE_SEQ", String(next));
    return "INV-" + String(next).padStart(6, "0");
  } finally {
    lock.releaseLock();
  }
}

/* ===================== SMALL UTILS ===================== */
function nowEST_() {
  return new Date(Utilities.formatDate(new Date(), CFG.TIMEZONE, "yyyy/MM/dd HH:mm:ss"));
}
function fmt_(d, f) {
  return Utilities.formatDate(d instanceof Date ? d : new Date(d), CFG.TIMEZONE, f);
}
function toNumber_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return isFinite(v) ? v : 0;
  const cleaned = v.toString().trim().replace(/[^0-9.\-]/g, "");
  const n = parseFloat(cleaned);
  return isFinite(n) ? n : 0;
}
function toTaxRate_(v) {
  if (v === null || v === undefined || v === "") return CFG.DEFAULT_TAX_RATE;
  if (typeof v === "number") return v > 1 ? v / 100 : v;
  const s = v.toString().trim();
  if (s.includes("%")) return toNumber_(s) / 100;
  const n = toNumber_(s);
  return n > 1 ? n / 100 : (n || CFG.DEFAULT_TAX_RATE);
}
function toYes_(s) {
  const v = (s || "").toString().trim().toLowerCase();
  return ["yes","y","true","1"].includes(v);
}
function safeName_(s) {
  const t = (s || "Guest").toString().replace(/[^a-zA-Z0-9 ]/g, "").trim();
  return (t.replace(/\s+/g, "_").substring(0, 20)) || "Guest";
}

/**
 * Normalize room number for comparison (handles string/number conversion, whitespace, leading zeros)
 * @param {*} room - Room number (can be string or number)
 * @returns {string} - Normalized room number
 */
function normalizeRoomNumber_(room) {
  if (!room) return "";
  return String(room).trim().replace(/^0+/, '');
}

/**
 * Check if a row has checkout status (checks both checkbox and CheckOutTime)
 * @param {Sheet} sh - The sheet object
 * @param {Array} rowData - Row data array
 * @param {number} rowNum - Row number
 * @param {object} C - Column indices
 * @returns {boolean} - True if row is checked out
 */
function isRowCheckedOut_(sh, rowData, rowNum, C) {
  if (!C.checkOut) return false;
  
  // Check checkbox value
  const checkOutValue = rowData[C.checkOut - 1];
  let checkOut = toYes_(checkOutValue);
  
  // Also check cell directly
  if (!checkOut) {
    const checkoutCell = sh.getRange(rowNum, C.checkOut).getValue();
    checkOut = toYes_(checkoutCell);
  }
  
  // Also check CheckOutTime as additional indicator
  if (!checkOut && C.checkOutTime) {
    const checkOutTime = rowData[C.checkOutTime - 1];
    if (checkOutTime) {
      checkOut = true;
    }
  }
  
  return checkOut;
}

/**
 * Check if a row has check-in status (checks both checkbox and CheckInTime)
 * @param {Sheet} sh - The sheet object
 * @param {Array} rowData - Row data array
 * @param {number} rowNum - Row number
 * @param {object} C - Column indices
 * @returns {boolean} - True if row is checked in
 */
function isRowCheckedIn_(sh, rowData, rowNum, C) {
  if (!C.checkIn) return false;
  
  // Check checkbox value
  let checkIn = C.checkIn ? toYes_(rowData[C.checkIn - 1]) : false;
  
  // Also check cell directly
  if (!checkIn) {
    const checkInCell = sh.getRange(rowNum, C.checkIn).getValue();
    checkIn = toYes_(checkInCell);
  }
  
  // Also check CheckInTime as additional indicator
  if (!checkIn && C.checkInTime) {
    const checkInTime = rowData[C.checkInTime - 1];
    if (checkInTime) {
      checkIn = true;
    }
  }
  
  return checkIn;
}

/* ===================== HOUSEKEEPING AUTOMATION ===================== */
/**
 * Check if a room is currently occupied by a newer check-in
 * @param {Sheet} sh - The sheet object
 * @param {string} room - Room number
 * @param {number} currentRow - Current row number
 * @param {object} C - Column indices
 * @returns {boolean} - True if room has a newer check-in that hasn't checked out
 */
function isRoomCurrentlyOccupied_(sh, room, currentRow, C) {
  if (!room || !C.room || !C.checkIn || !C.checkOut) return false;
  
  const data = sh.getDataRange().getValues();
  const normalizedRoom = normalizeRoomNumber_(room);
  
  // Check all rows after the current row for the same room
  for (let i = currentRow; i < data.length; i++) {
    const rowNum = i + 1;
    const rowRoom = data[i][C.room - 1];
    if (!rowRoom) continue;
    
    const normalizedRowRoom = normalizeRoomNumber_(rowRoom);
    
    if (normalizedRoom === normalizedRowRoom) {
      const checkIn = isRowCheckedIn_(sh, data[i], rowNum, C);
      const checkOut = isRowCheckedOut_(sh, data[i], rowNum, C);
      
      // If there's a check-in without a check-out, room is occupied
      if (checkIn && !checkOut) {
        return true;
      }
    }
  }
  
  return false;
}

/**
 * Clear "Cleaned - ReadyFor Rent" status from older rows for the same room
 * @param {Sheet} sh - The sheet object
 * @param {string} room - Room number
 * @param {number} currentRow - Current row number (new check-in)
 * @param {object} C - Column indices
 */
function clearOldCleanedStatus_(sh, room, currentRow, C) {
  if (!room || !C.room || !C.hkStatus) return;
  
  const data = sh.getDataRange().getValues();
  const normalizedTargetRoom = normalizeRoomNumber_(room);
  
  // Check all rows before the current row for the same room
  for (let i = 1; i < currentRow - 1; i++) {
    const rowNum = i + 1;
    const rowRoom = data[i][C.room - 1];
    if (!rowRoom) continue;
    
    const normalizedRowRoom = normalizeRoomNumber_(rowRoom);
    
    if (normalizedTargetRoom === normalizedRowRoom) {
      const hkStatus = C.hkStatus ? (data[i][C.hkStatus - 1] || "").toString() : "";
      
      // If status is "Cleaned - ReadyFor Rent", clear it
      if (hkStatus === CFG.HK_DONE_TEXT) {
        sh.getRange(rowNum, C.hkStatus).setValue("").setBackground("");
        log_("Cleared old cleaned status", { room, oldRow: rowNum, newRow: currentRow });
      }
    }
  }
}

/**
 * Apply gray styling to a row (for old/inactive data)
 * @param {Sheet} sh - The sheet object
 * @param {number} rowNum - Row number to style
 * @param {object} C - Column indices
 */
function applyOldRowStyling_(sh, rowNum, C) {
  const startCol = C.date || 1;
  const endCol = Math.max(CFG.OLD_DATA_END_COLUMN, sh.getLastColumn());
  const rowRange = sh.getRange(rowNum, startCol, 1, endCol - startCol + 1);
  
  rowRange.setBackground(CFG.COLOR.OLD_DATA);
  rowRange.setFontColor(CFG.COLOR.OLD_DATA_TEXT);
}

/**
 * Gray out entire rows for old data (rows that have been checked out
 * and the room has been re-rented to a new guest)
 * @param {Sheet} sh - The sheet object
 * @param {string} room - Room number (optional, if provided only processes this room)
 * @param {number} newCheckInRow - Row number of new check-in (optional)
 * @param {object} C - Column indices
 * @returns {number} - Count of rows grayed out
 */
function grayOutOldRows_(sh, room, newCheckInRow, C) {
  if (!C.date || !C.checkOut || !C.room || !C.checkIn) return;
  
  const data = sh.getDataRange().getValues();
  const lastRow = sh.getLastRow();
  let grayedOutCount = 0;
  
  // If a specific room and row are provided, only process that room's old rows
  if (room && newCheckInRow) {
    // Process only older rows for this specific room
    const normalizedTargetRoom = normalizeRoomNumber_(room);
    
    for (let i = 1; i < newCheckInRow - 1; i++) {
      const rowNum = i + 1;
      const rowRoom = data[i][C.room - 1];
      if (!rowRoom) continue;
      
      const normalizedRowRoom = normalizeRoomNumber_(rowRoom);
      
      if (normalizedTargetRoom !== normalizedRowRoom) continue;
      
      // Check if row is checked out
      const checkOut = isRowCheckedOut_(sh, data[i], rowNum, C);
      
      if (checkOut) {
        try {
          applyOldRowStyling_(sh, rowNum, C);
          grayedOutCount++;
          log_("Grayed out old row", { room, row: rowNum });
        } catch (rangeError) {
          log_("Error graying out row", { room, row: rowNum, error: rangeError.toString() });
        }
      }
    }
  } else {
    // Process all rows to find and gray out old data
    for (let i = 1; i < data.length; i++) {
      const rowNum = i + 1;
      const rowRoom = data[i][C.room - 1];
      
      // Skip if no room number
      if (!rowRoom) continue;
      
      // Check if row is checked out
      const checkOut = isRowCheckedOut_(sh, data[i], rowNum, C);
      
      // Skip if no checkout (but log for debugging)
      if (!checkOut) {
        log_("Skipped row (no checkout)", { row: rowNum, room: rowRoom });
        continue;
      }
      
      // Check if there's a newer row for the same room with a check-in
      // (regardless of whether that newer check-in has also checked out)
      let hasNewerCheckIn = false;
      const normalizedRoom = normalizeRoomNumber_(rowRoom);
      
      for (let j = i + 1; j < data.length; j++) {
        const newerRoom = data[j][C.room - 1];
        if (!newerRoom) continue;
        
        const normalizedNewerRoom = normalizeRoomNumber_(newerRoom);
        
        // Compare room numbers (should match after normalization)
        if (normalizedRoom === normalizedNewerRoom) {
          const checkInValue = isRowCheckedIn_(sh, data[j], j + 1, C);
          
          if (checkInValue) {
            hasNewerCheckIn = true;
            log_("Found newer check-in", { oldRow: rowNum, newRow: j + 1, room: normalizedRoom });
            break;
          }
        }
      }
      
      // If there's a newer check-in, gray out this old row
      if (hasNewerCheckIn) {
        try {
          applyOldRowStyling_(sh, rowNum, C);
          grayedOutCount++;
          log_("Grayed out old row", { room: rowRoom, row: rowNum });
        } catch (rangeError) {
          log_("Error graying out row", { room: rowRoom, row: rowNum, error: rangeError.toString() });
        }
      } else {
        // Debug: log rows that have checkout but no newer check-in
        log_("Skipped row (no newer check-in)", { room: rowRoom, row: rowNum, checkOut: checkOut });
      }
    }
  }
  
  log_("Grayed out old rows completed", { count: grayedOutCount });
  return grayedOutCount;
}

function handleHKDone_(sh, row, C) {
  try {
    if (!C.hkDone) return;
    
    const room = sh.getRange(row, C.room).getValue();
    
    // Check if room is already occupied by a newer check-in
    // If so, don't set "Cleaned - ReadyFor Rent" because the room is already rented
    const isRoomOccupied = isRoomCurrentlyOccupied_(sh, room, row, C);
    
    // Set cleaned time if not already set
    if (C.cleanedTime && !sh.getRange(row, C.cleanedTime).getValue()) {
      sh.getRange(row, C.cleanedTime).setValue(nowEST_());
    }
    
    // Update HK Status only if room is not currently occupied
    if (C.hkStatus) {
      if (isRoomOccupied) {
        // Room is already occupied, clear HK Status to avoid confusion
        sh.getRange(row, C.hkStatus).setValue("").setBackground("");
      } else {
        // Room is available, set to "Cleaned - ReadyFor Rent"
        sh.getRange(row, C.hkStatus).setValue(CFG.HK_DONE_TEXT).setBackground(CFG.COLOR.HK_DONE);
      }
    }
    
    // Color code HK Done cell
    sh.getRange(row, C.hkDone).setBackground(CFG.COLOR.HK_DONE);
    
    // Update Daily Dashboard (non-blocking)
    try {
      updateDailyDashboard_();
    } catch (error) {
      log_("Error updating dashboard on HK Done", { error: error.toString() });
    }
    
    log_("HK Done processed", { room: room, row, isRoomOccupied });
  } catch (error) {
    log_("Error in handleHKDone_", { error: error.toString(), row });
  }
}

/* ===================== PAYMENT TYPE HANDLING ===================== */
function handlePaymentTypeChange_(sh, row, C, paymentType) {
  try {
    if (!C.paymentType) return;
    
    const normalized = paymentType.trim();
    let color = CFG.COLOR.PAYMENT_OTHER;
    
    if (normalized.toLowerCase().includes("cash")) {
      color = CFG.COLOR.PAYMENT_CASH;
    } else if (normalized.toLowerCase().includes("card") || normalized.toLowerCase().includes("credit") || normalized.toLowerCase().includes("debit")) {
      color = CFG.COLOR.PAYMENT_CARD;
    }
    
    sh.getRange(row, C.paymentType).setBackground(color);
    
    // Validate payment type if needed
    if (CFG.PAYMENT_TYPES.length > 0 && !CFG.PAYMENT_TYPES.some(pt => normalized.toLowerCase().includes(pt.toLowerCase()))) {
      // Payment type not in standard list, but allow it
      log_("Non-standard payment type", { paymentType: normalized, row });
    }
    
  } catch (error) {
    log_("Error in handlePaymentTypeChange_", { error: error.toString(), row });
  }
}

/* ===================== REPORTING ===================== */
function menuDailyReport() {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
    if (!sh) {
      SpreadsheetApp.getUi().alert("Sheet not found: " + CFG.SHEET);
      return;
    }
    
    const today = nowEST_();
    const report = generateDailyReport_(sh, today);
    showReportDialog_("Daily Report - " + fmt_(today, "MM/dd/yyyy"), report);
  } catch (error) {
    log_("Error in menuDailyReport", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error generating report: " + error.toString());
  }
}

function menuWeeklyReport() {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
    if (!sh) {
      SpreadsheetApp.getUi().alert("Sheet not found: " + CFG.SHEET);
      return;
    }
    
    const today = nowEST_();
    const report = generateWeeklyReport_(sh, today);
    showReportDialog_("Weekly Report - Week of " + fmt_(today, "MM/dd/yyyy"), report);
  } catch (error) {
    log_("Error in menuWeeklyReport", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error generating report: " + error.toString());
  }
}

function menuMonthlyReport() {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
    if (!sh) {
      SpreadsheetApp.getUi().alert("Sheet not found: " + CFG.SHEET);
      return;
    }
    
    const today = nowEST_();
    const report = generateMonthlyReport_(sh, today);
    showReportDialog_("Monthly Report - " + fmt_(today, "MMMM yyyy"), report);
  } catch (error) {
    log_("Error in menuMonthlyReport", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error generating report: " + error.toString());
  }
}

function menuRoomOccupancyReport() {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
    if (!sh) {
      SpreadsheetApp.getUi().alert("Sheet not found: " + CFG.SHEET);
      return;
    }
    
    const report = generateRoomOccupancyReport_(sh);
    showReportDialog_("Room Occupancy Report", report);
  } catch (error) {
    log_("Error in menuRoomOccupancyReport", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error generating report: " + error.toString());
  }
}

function generateDailyReport_(sh, date) {
  const C = cols_(sh);
  const dateStr = fmt_(date, "MM/dd/yyyy");
  const data = sh.getDataRange().getValues();
  
  let checkIns = 0, checkOuts = 0, revenue = 0, roomsOccupied = 0;
  const rooms = new Set();
  
  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const rowDate = C.date ? (data[i][C.date - 1] ? fmt_(data[i][C.date - 1], "MM/dd/yyyy") : "") : "";
    const checkIn = C.checkIn ? toYes_(data[i][C.checkIn - 1]) : false;
    const checkOut = C.checkOut ? toYes_(data[i][C.checkOut - 1]) : false;
    const total = C.total ? toNumber_(data[i][C.total - 1]) : 0;
    const room = C.room ? data[i][C.room - 1] : "";
    
    if (rowDate === dateStr || (!rowDate && (checkIn || checkOut))) {
      if (checkIn) {
        checkIns++;
        if (room) rooms.add(room);
      }
      if (checkOut) {
        checkOuts++;
        revenue += total;
      }
    }
    
    // Count currently occupied rooms
    if (checkIn && !checkOut && room) {
      roomsOccupied++;
    }
  }
  
  return {
    date: dateStr,
    checkIns: checkIns,
    checkOuts: checkOuts,
    revenue: revenue,
    roomsOccupied: roomsOccupied,
    occupancyRate: rooms.size > 0 ? ((roomsOccupied / rooms.size) * 100).toFixed(1) + "%" : "N/A"
  };
}

function generateWeeklyReport_(sh, date) {
  const startOfWeek = new Date(date);
  startOfWeek.setDate(date.getDate() - date.getDay()); // Sunday
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 6);
  
  const C = cols_(sh);
  const data = sh.getDataRange().getValues();
  
  let totalRevenue = 0, totalCheckIns = 0, totalCheckOuts = 0;
  const dailyStats = {};
  
  for (let i = 1; i < data.length; i++) {
    const rowDate = C.date ? data[i][C.date - 1] : null;
    if (!rowDate) continue;
    
    const checkIn = C.checkIn ? toYes_(data[i][C.checkIn - 1]) : false;
    const checkOut = C.checkOut ? toYes_(data[i][C.checkOut - 1]) : false;
    const total = C.total ? toNumber_(data[i][C.total - 1]) : 0;
    const dateObj = rowDate instanceof Date ? rowDate : new Date(rowDate);
    
    if (dateObj >= startOfWeek && dateObj <= endOfWeek) {
      const dayStr = fmt_(dateObj, "MM/dd/yyyy");
      if (!dailyStats[dayStr]) {
        dailyStats[dayStr] = { checkIns: 0, checkOuts: 0, revenue: 0 };
      }
      
      if (checkIn) {
        dailyStats[dayStr].checkIns++;
        totalCheckIns++;
      }
      if (checkOut) {
        dailyStats[dayStr].checkOuts++;
        dailyStats[dayStr].revenue += total;
        totalCheckOuts++;
        totalRevenue += total;
      }
    }
  }
  
  return {
    period: fmt_(startOfWeek, "MM/dd/yyyy") + " to " + fmt_(endOfWeek, "MM/dd/yyyy"),
    totalRevenue: totalRevenue,
    totalCheckIns: totalCheckIns,
    totalCheckOuts: totalCheckOuts,
    dailyStats: dailyStats
  };
}

function generateMonthlyReport_(sh, date) {
  const year = date.getFullYear();
  const month = date.getMonth();
  const startOfMonth = new Date(year, month, 1);
  const endOfMonth = new Date(year, month + 1, 0);
  
  const C = cols_(sh);
  const data = sh.getDataRange().getValues();
  
  let totalRevenue = 0, totalCheckIns = 0, totalCheckOuts = 0;
  const paymentTypes = {};
  
  for (let i = 1; i < data.length; i++) {
    const rowDate = C.date ? data[i][C.date - 1] : null;
    if (!rowDate) continue;
    
    const dateObj = rowDate instanceof Date ? rowDate : new Date(rowDate);
    if (dateObj.getMonth() === month && dateObj.getFullYear() === year) {
      const checkIn = C.checkIn ? toYes_(data[i][C.checkIn - 1]) : false;
      const checkOut = C.checkOut ? toYes_(data[i][C.checkOut - 1]) : false;
      const total = C.total ? toNumber_(data[i][C.total - 1]) : 0;
      const paymentType = C.paymentType ? (data[i][C.paymentType - 1] || "Unknown") : "Unknown";
      
      if (checkIn) totalCheckIns++;
      if (checkOut) {
        totalCheckOuts++;
        totalRevenue += total;
        
        if (!paymentTypes[paymentType]) paymentTypes[paymentType] = 0;
        paymentTypes[paymentType] += total;
      }
    }
  }
  
  return {
    period: fmt_(startOfMonth, "MMMM yyyy"),
    totalRevenue: totalRevenue,
    totalCheckIns: totalCheckIns,
    totalCheckOuts: totalCheckOuts,
    paymentTypes: paymentTypes,
    averageRevenue: totalCheckOuts > 0 ? (totalRevenue / totalCheckOuts).toFixed(2) : 0
  };
}

function generateRoomOccupancyReport_(sh) {
  const C = cols_(sh);
  const data = sh.getDataRange().getValues();
  
  const roomStats = {};
  
  for (let i = 1; i < data.length; i++) {
    const room = C.room ? data[i][C.room - 1] : "";
    if (!room) continue;
    
    if (!roomStats[room]) {
      roomStats[room] = {
        checkIns: 0,
        checkOuts: 0,
        revenue: 0,
        nights: 0
      };
    }
    
    const checkIn = C.checkIn ? toYes_(data[i][C.checkIn - 1]) : false;
    const checkOut = C.checkOut ? toYes_(data[i][C.checkOut - 1]) : false;
    const total = C.total ? toNumber_(data[i][C.total - 1]) : 0;
    const nights = C.nights ? toNumber_(data[i][C.nights - 1]) : 0;
    
    if (checkIn) roomStats[room].checkIns++;
    if (checkOut) {
      roomStats[room].checkOuts++;
      roomStats[room].revenue += total;
      roomStats[room].nights += nights;
    }
  }
  
  return roomStats;
}

function showReportDialog_(title, data) {
  let html = `<html><body style="font-family:Arial; padding:20px;"><h2>${title}</h2>`;
  
  if (data.date) {
    // Daily report
    html += `<p><b>Date:</b> ${data.date}</p>`;
    html += `<p><b>Check-ins:</b> ${data.checkIns}</p>`;
    html += `<p><b>Check-outs:</b> ${data.checkOuts}</p>`;
    html += `<p><b>Revenue:</b> $${data.revenue.toFixed(2)}</p>`;
    html += `<p><b>Rooms Occupied:</b> ${data.roomsOccupied}</p>`;
    html += `<p><b>Occupancy Rate:</b> ${data.occupancyRate}</p>`;
  } else if (data.period) {
    // Weekly report
    html += `<p><b>Period:</b> ${data.period}</p>`;
    html += `<p><b>Total Check-ins:</b> ${data.totalCheckIns}</p>`;
    html += `<p><b>Total Check-outs:</b> ${data.totalCheckOuts}</p>`;
    html += `<p><b>Total Revenue:</b> $${data.totalRevenue.toFixed(2)}</p>`;
    html += `<h3>Daily Breakdown:</h3><ul>`;
    for (const [day, stats] of Object.entries(data.dailyStats)) {
      html += `<li>${day}: ${stats.checkIns} check-ins, ${stats.checkOuts} check-outs, $${stats.revenue.toFixed(2)}</li>`;
    }
    html += `</ul>`;
  } else if (data.paymentTypes) {
    // Monthly report
    html += `<p><b>Period:</b> ${data.period}</p>`;
    html += `<p><b>Total Check-ins:</b> ${data.totalCheckIns}</p>`;
    html += `<p><b>Total Check-outs:</b> ${data.totalCheckOuts}</p>`;
    html += `<p><b>Total Revenue:</b> $${data.totalRevenue.toFixed(2)}</p>`;
    html += `<p><b>Average Revenue per Check-out:</b> $${data.averageRevenue}</p>`;
    html += `<h3>Revenue by Payment Type:</h3><ul>`;
    for (const [type, amount] of Object.entries(data.paymentTypes)) {
      html += `<li>${type}: $${amount.toFixed(2)}</li>`;
    }
    html += `</ul>`;
  } else {
    // Room occupancy report
    html += `<h3>Room Statistics:</h3><table border="1" cellpadding="5" style="border-collapse:collapse;"><tr><th>Room</th><th>Check-ins</th><th>Check-outs</th><th>Revenue</th><th>Total Nights</th></tr>`;
    for (const [room, stats] of Object.entries(data)) {
      html += `<tr><td>${room}</td><td>${stats.checkIns}</td><td>${stats.checkOuts}</td><td>$${stats.revenue.toFixed(2)}</td><td>${stats.nights}</td></tr>`;
    }
    html += `</table>`;
  }
  
  html += `</body></html>`;
  
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(600).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

/* ===================== DASHBOARD ===================== */
function menuRoomAvailability() {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
    if (!sh) {
      SpreadsheetApp.getUi().alert("Sheet not found: " + CFG.SHEET);
      return;
    }
    
    const availability = getRoomAvailability_(sh);
    showAvailabilityDialog_(availability);
  } catch (error) {
    log_("Error in menuRoomAvailability", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

function menuTodaysCheckins() {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
    if (!sh) {
      SpreadsheetApp.getUi().alert("Sheet not found: " + CFG.SHEET);
      return;
    }
    
    const today = fmt_(nowEST_(), "MM/dd/yyyy");
    const checkins = getTodaysCheckins_(sh, today);
    showListDialog_("Today's Check-ins (" + today + ")", checkins);
  } catch (error) {
    log_("Error in menuTodaysCheckins", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

function menuTodaysCheckouts() {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
    if (!sh) {
      SpreadsheetApp.getUi().alert("Sheet not found: " + CFG.SHEET);
      return;
    }
    
    const today = fmt_(nowEST_(), "MM/dd/yyyy");
    const checkouts = getTodaysCheckouts_(sh, today);
    showListDialog_("Today's Check-outs (" + today + ")", checkouts);
  } catch (error) {
    log_("Error in menuTodaysCheckouts", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

function menuPendingHousekeeping() {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
    if (!sh) {
      SpreadsheetApp.getUi().alert("Sheet not found: " + CFG.SHEET);
      return;
    }
    
    const pending = getPendingHousekeeping_(sh);
    showPendingHousekeepingReport_(pending);
  } catch (error) {
    log_("Error in menuPendingHousekeeping", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

function getRoomAvailability_(sh) {
  const C = cols_(sh);
  const data = sh.getDataRange().getValues();
  const rooms = {};
  
  for (let i = 1; i < data.length; i++) {
    const room = C.room ? data[i][C.room - 1] : "";
    if (!room) continue;
    
    if (!rooms[room]) {
      rooms[room] = { status: "Available", guest: "", checkInTime: "", checkOutTime: "" };
    }
    
    const checkIn = C.checkIn ? toYes_(data[i][C.checkIn - 1]) : false;
    const checkOut = C.checkOut ? toYes_(data[i][C.checkOut - 1]) : false;
    const guest = C.guest ? data[i][C.guest - 1] : "";
    const checkInTime = C.checkInTime ? data[i][C.checkInTime - 1] : "";
    const checkOutTime = C.checkOutTime ? data[i][C.checkOutTime - 1] : "";
    const hkDone = C.hkDone ? toYes_(data[i][C.hkDone - 1]) : false;
    
    if (checkIn && !checkOut) {
      rooms[room].status = "Occupied";
      rooms[room].guest = guest;
      rooms[room].checkInTime = checkInTime ? fmt_(checkInTime, "MM/dd/yyyy hh:mm a") : "";
    } else if (checkOut && !hkDone) {
      rooms[room].status = "Ready for Cleaning";
      rooms[room].guest = guest;
      rooms[room].checkOutTime = checkOutTime ? fmt_(checkOutTime, "MM/dd/yyyy hh:mm a") : "";
    } else if (checkOut && hkDone) {
      rooms[room].status = "Available";
    }
  }
  
  return rooms;
}

function getTodaysCheckins_(sh, today) {
  const C = cols_(sh);
  const data = sh.getDataRange().getValues();
  const checkins = [];
  
  for (let i = 1; i < data.length; i++) {
    const rowDate = C.date ? (data[i][C.date - 1] ? fmt_(data[i][C.date - 1], "MM/dd/yyyy") : "") : "";
    const checkIn = C.checkIn ? toYes_(data[i][C.checkIn - 1]) : false;
    const checkInTime = C.checkInTime ? data[i][C.checkInTime - 1] : "";
    
    if (checkIn && (rowDate === today || (checkInTime && fmt_(checkInTime, "MM/dd/yyyy") === today))) {
      checkins.push({
        room: C.room ? data[i][C.room - 1] : "",
        guest: C.guest ? data[i][C.guest - 1] : "",
        time: checkInTime ? fmt_(checkInTime, "hh:mm a") : ""
      });
    }
  }
  
  return checkins;
}

function getTodaysCheckouts_(sh, today) {
  const C = cols_(sh);
  const data = sh.getDataRange().getValues();
  const checkouts = [];
  
  for (let i = 1; i < data.length; i++) {
    const rowDate = C.date ? (data[i][C.date - 1] ? fmt_(data[i][C.date - 1], "MM/dd/yyyy") : "") : "";
    const checkOut = C.checkOut ? toYes_(data[i][C.checkOut - 1]) : false;
    const checkOutTime = C.checkOutTime ? data[i][C.checkOutTime - 1] : "";
    
    if (checkOut && (rowDate === today || (checkOutTime && fmt_(checkOutTime, "MM/dd/yyyy") === today))) {
      checkouts.push({
        room: C.room ? data[i][C.room - 1] : "",
        guest: C.guest ? data[i][C.guest - 1] : "",
        total: C.total ? toNumber_(data[i][C.total - 1]) : 0,
        time: checkOutTime ? fmt_(checkOutTime, "hh:mm a") : ""
      });
    }
  }
  
  return checkouts;
}

function getPendingHousekeeping_(sh) {
  const C = cols_(sh);
  const data = sh.getDataRange().getValues();
  const pending = [];
  
  for (let i = 1; i < data.length; i++) {
    const checkOut = C.checkOut ? toYes_(data[i][C.checkOut - 1]) : false;
    const hkDone = C.hkDone ? toYes_(data[i][C.hkDone - 1]) : false;
    
    if (checkOut && !hkDone) {
      pending.push({
        room: C.room ? data[i][C.room - 1] : "",
        guest: C.guest ? data[i][C.guest - 1] : "",
        checkOutTime: C.checkOutTime ? (data[i][C.checkOutTime - 1] ? fmt_(data[i][C.checkOutTime - 1], "MM/dd/yyyy hh:mm a") : "") : ""
      });
    }
  }
  
  return pending;
}

function showAvailabilityDialog_(availability) {
  let html = `<html><body style="font-family:Arial; padding:20px;"><h2>Room Availability</h2>`;
  html += `<table border="1" cellpadding="5" style="border-collapse:collapse; width:100%;"><tr><th>Room</th><th>Status</th><th>Guest</th><th>Time</th></tr>`;
  
  const sortedRooms = Object.keys(availability).sort();
  for (const room of sortedRooms) {
    const info = availability[room];
    html += `<tr><td>${room}</td><td>${info.status}</td><td>${info.guest || "-"}</td><td>${info.checkInTime || info.checkOutTime || "-"}</td></tr>`;
  }
  
  html += `</table></body></html>`;
  
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(700).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Room Availability");
}

function showListDialog_(title, items) {
  let html = `<html><body style="font-family:Arial; padding:20px;"><h2>${title}</h2>`;
  
  if (items.length === 0) {
    html += `<p>No items found.</p>`;
  } else {
    html += `<table border="1" cellpadding="5" style="border-collapse:collapse; width:100%;">`;
    
    // Determine columns from first item
    const keys = Object.keys(items[0] || {});
    html += `<tr>`;
    for (const key of keys) {
      html += `<th>${key.charAt(0).toUpperCase() + key.slice(1).replace(/([A-Z])/g, ' $1')}</th>`;
    }
    html += `</tr>`;
    
    for (const item of items) {
      html += `<tr>`;
      for (const key of keys) {
        const value = item[key];
        html += `<td>${value || "-"}</td>`;
      }
      html += `</tr>`;
    }
    
    html += `</table>`;
  }
  
  html += `</body></html>`;
  
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(700).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

/* ===================== GUEST HISTORY ===================== */
function menuGuestHistory() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Guest History Lookup", "Enter guest name (partial match OK):", ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() !== ui.Button.OK) return;
    
    const searchTerm = response.getResponseText().trim();
    if (!searchTerm) {
      ui.alert("Please enter a guest name.");
      return;
    }
    
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
    if (!sh) {
      ui.alert("Sheet not found: " + CFG.SHEET);
      return;
    }
    
    const history = getGuestHistory_(sh, searchTerm);
    showGuestHistoryDialog_(searchTerm, history);
  } catch (error) {
    log_("Error in menuGuestHistory", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

function getGuestHistory_(sh, searchTerm) {
  const C = cols_(sh);
  const data = sh.getDataRange().getValues();
  const history = [];
  const searchLower = searchTerm.toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    const guest = C.guest ? (data[i][C.guest - 1] || "").toString() : "";
    if (guest.toLowerCase().includes(searchLower)) {
      history.push({
        date: C.date ? (data[i][C.date - 1] ? fmt_(data[i][C.date - 1], "MM/dd/yyyy") : "") : "",
        room: C.room ? data[i][C.room - 1] : "",
        guest: guest,
        nights: C.nights ? toNumber_(data[i][C.nights - 1]) : 0,
        total: C.total ? toNumber_(data[i][C.total - 1]) : 0,
        checkIn: C.checkIn ? toYes_(data[i][C.checkIn - 1]) : false,
        checkOut: C.checkOut ? toYes_(data[i][C.checkOut - 1]) : false,
        invoiceNo: C.invoiceNo ? data[i][C.invoiceNo - 1] : ""
      });
    }
  }
  
  return history;
}

function showGuestHistoryDialog_(searchTerm, history) {
  let html = `<html><body style="font-family:Arial; padding:20px;"><h2>Guest History: "${searchTerm}"</h2>`;
  
  if (history.length === 0) {
    html += `<p>No history found for "${searchTerm}".</p>`;
  } else {
    html += `<p>Found ${history.length} record(s)</p>`;
    html += `<table border="1" cellpadding="5" style="border-collapse:collapse; width:100%;">`;
    html += `<tr><th>Date</th><th>Room</th><th>Guest</th><th>Nights</th><th>Total</th><th>Status</th><th>Invoice</th></tr>`;
    
    for (const record of history) {
      let status = "";
      if (record.checkIn && record.checkOut) status = "Checked Out";
      else if (record.checkIn) status = "Checked In";
      else status = "Pending";
      
      html += `<tr>`;
      html += `<td>${record.date || "-"}</td>`;
      html += `<td>${record.room || "-"}</td>`;
      html += `<td>${record.guest || "-"}</td>`;
      html += `<td>${record.nights}</td>`;
      html += `<td>$${record.total.toFixed(2)}</td>`;
      html += `<td>${status}</td>`;
      html += `<td>${record.invoiceNo || "-"}</td>`;
      html += `</tr>`;
    }
    
    html += `</table>`;
    
    // Summary
    const totalRevenue = history.reduce((sum, r) => sum + r.total, 0);
    const totalNights = history.reduce((sum, r) => sum + r.nights, 0);
    html += `<p><b>Total Revenue:</b> $${totalRevenue.toFixed(2)} | <b>Total Nights:</b> ${totalNights}</p>`;
  }
  
  html += `</body></html>`;
  
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(900).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Guest History");
}

/* ===================== BULK OPERATIONS ===================== */
function menuBulkGenerateInvoices() {
  try {
    const sh = SpreadsheetApp.getActiveSheet();
    if (sh.getName() !== CFG.SHEET) {
      SpreadsheetApp.getUi().alert("Run this from FrontDesk_Log.");
      return;
    }
    
    const range = sh.getActiveRange();
    if (!range) {
      SpreadsheetApp.getUi().alert("Please select rows first.");
      return;
    }
    
    const startRow = range.getRow();
    const endRow = range.getLastRow();
    
    if (startRow < 2) {
      SpreadsheetApp.getUi().alert("Please select data rows (not header).");
      return;
    }
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert("Bulk Generate Invoices", 
      `Generate invoices for ${endRow - startRow + 1} selected row(s)?`, 
      ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) return;
    
    const C = cols_(sh);
    let count = 0;
    
    for (let row = startRow; row <= endRow; row++) {
      try {
        updateQuoteForRow_(sh, row, C);
        generateInvoiceForRow_(sh, row, C, true);
        count++;
      } catch (error) {
        log_("Error generating invoice for row", { row, error: error.toString() });
      }
    }
    
    ui.alert(`Generated ${count} invoice(s).`);
  } catch (error) {
    log_("Error in menuBulkGenerateInvoices", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

function menuBulkUpdateTaxRate() {
  try {
    const sh = SpreadsheetApp.getActiveSheet();
    if (sh.getName() !== CFG.SHEET) {
      SpreadsheetApp.getUi().alert("Run this from FrontDesk_Log.");
      return;
    }
    
    const range = sh.getActiveRange();
    if (!range) {
      SpreadsheetApp.getUi().alert("Please select rows first.");
      return;
    }
    
    const startRow = range.getRow();
    const endRow = range.getLastRow();
    
    if (startRow < 2) {
      SpreadsheetApp.getUi().alert("Please select data rows (not header).");
      return;
    }
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Bulk Update Tax Rate", 
      `Enter new tax rate (e.g., 0.13 or 13%):`, ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() !== ui.Button.OK) return;
    
    const taxRateInput = response.getResponseText().trim();
    const taxRate = toTaxRate_(taxRateInput);
    
    if (taxRate <= 0 || taxRate > 1) {
      ui.alert("Invalid tax rate. Please enter a value between 0 and 1 (e.g., 0.13) or percentage (e.g., 13%).");
      return;
    }
    
    const C = cols_(sh);
    if (!C.taxRate) {
      ui.alert("Tax Rate column not found.");
      return;
    }
    
    let count = 0;
    for (let row = startRow; row <= endRow; row++) {
      sh.getRange(row, C.taxRate).setValue(taxRate);
      updateQuoteForRow_(sh, row, C);
      count++;
    }
    
    ui.alert(`Updated tax rate for ${count} row(s).`);
  } catch (error) {
    log_("Error in menuBulkUpdateTaxRate", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

function menuBulkCheckout() {
  try {
    const sh = SpreadsheetApp.getActiveSheet();
    if (sh.getName() !== CFG.SHEET) {
      SpreadsheetApp.getUi().alert("Run this from FrontDesk_Log.");
      return;
    }
    
    const range = sh.getActiveRange();
    if (!range) {
      SpreadsheetApp.getUi().alert("Please select rows first.");
      return;
    }
    
    const startRow = range.getRow();
    const endRow = range.getLastRow();
    
    if (startRow < 2) {
      SpreadsheetApp.getUi().alert("Please select data rows (not header).");
      return;
    }
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert("Bulk Check-out", 
      `Mark ${endRow - startRow + 1} selected row(s) as checked out?`, 
      ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) return;
    
    const C = cols_(sh);
    if (!C.checkOut) {
      ui.alert("CheckOut column not found.");
      return;
    }
    
    let count = 0;
    for (let row = startRow; row <= endRow; row++) {
      const checkIn = C.checkIn ? toYes_(sh.getRange(row, C.checkIn).getValue()) : false;
      const checkOut = C.checkOut ? toYes_(sh.getRange(row, C.checkOut).getValue()) : false;
      
      if (checkIn && !checkOut) {
        sh.getRange(row, C.checkOut).setValue("Yes");
        // Trigger the edit handler logic
        if (!sh.getRange(row, C.checkOutTime).getValue()) {
          sh.getRange(row, C.checkOutTime).setValue(nowEST_());
        }
        sh.getRange(row, C.checkOut).setBackground(CFG.COLOR.CHECKOUT_YES);
        
        if (C.hkStatus) {
          sh.getRange(row, C.hkStatus).setValue(CFG.HK_READY_TEXT).setBackground(CFG.COLOR.HK_READY);
        }
        
        updateQuoteForRow_(sh, row, C);
        generateInvoiceForRow_(sh, row, C, false);
        count++;
      }
    }
    
    ui.alert(`Processed ${count} check-out(s).`);
  } catch (error) {
    log_("Error in menuBulkCheckout", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

function menuRefreshOldRowStyling() {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
    if (!sh) {
      SpreadsheetApp.getUi().alert("Sheet not found: " + CFG.SHEET);
      return;
    }
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert("Refresh Old Row Styling", 
      "This will gray out Date and CheckOut columns for old rows (rooms that have been re-rented).\n\nContinue?",
      ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) return;
    
    const C = cols_(sh);
    if (!C.date || !C.checkOut || !C.room || !C.checkIn) {
      ui.alert("Error: Required columns not found. Please ensure Date, CheckOut, Room #, and CheckIn columns exist.");
      return;
    }
    
    // Debug: Show column indices
    log_("Column indices", { 
      date: C.date, 
      checkOut: C.checkOut, 
      room: C.room, 
      checkIn: C.checkIn 
    });
    
    const count = grayOutOldRows_(sh, null, null, C);
    SpreadsheetApp.flush(); // Ensure changes are applied
    
    ui.alert(`Old row styling refreshed successfully!\n\nGrayed out ${count} row(s).`);
  } catch (error) {
    log_("Error in menuRefreshOldRowStyling", { error: error.toString(), stack: error.stack });
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

function menuSetupOnce() {
  try {
    setupOnce();
  } catch (error) {
    log_("Error in menuSetupOnce", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

/* ===================== DAILY DASHBOARD ===================== */
function menuUpdateDailyDashboard() {
  try {
    updateDailyDashboard_();
    SpreadsheetApp.getUi().alert("Daily Dashboard updated successfully!");
  } catch (error) {
    log_("Error in menuUpdateDailyDashboard", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error updating dashboard: " + error.toString());
  }
}

function updateDailyDashboard_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CFG.SHEET);
  const dashboardSheet = ss.getSheetByName("Daily_Dashboard");
  
  if (!logSheet) {
    throw new Error(`Sheet not found: ${CFG.SHEET}`);
  }
  
  if (!dashboardSheet) {
    throw new Error("Daily_Dashboard sheet not found. Please create it first.");
  }
  
  const today = nowEST_();
  const todayStr = fmt_(today, "MM/dd/yyyy");
  
  // Update report date
  dashboardSheet.getRange("B3").setValue(todayStr);
  
  // Get data from FrontDesk_Log
  const C = cols_(logSheet);
  const data = logSheet.getDataRange().getValues();
  
  let checkedInToday = 0;
  let checkedOutToday = 0;
  let readyForCleaningToday = 0;
  let cleanedToday = 0;
  let cashTotal = 0;
  let cardTotal = 0;
  let totalRevenue = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    
    // Get date - check both Date column and CheckInTime/CheckOutTime
    let rowDate = "";
    if (C.date) {
      const dateVal = data[i][C.date - 1];
      if (dateVal) {
        rowDate = fmt_(dateVal instanceof Date ? dateVal : new Date(dateVal), "MM/dd/yyyy");
      }
    }
    
    // If no date column or empty, check timestamps
    if (!rowDate) {
      if (C.checkInTime && data[i][C.checkInTime - 1]) {
        const checkInTime = data[i][C.checkInTime - 1];
        rowDate = fmt_(checkInTime instanceof Date ? checkInTime : new Date(checkInTime), "MM/dd/yyyy");
      } else if (C.checkOutTime && data[i][C.checkOutTime - 1]) {
        const checkOutTime = data[i][C.checkOutTime - 1];
        rowDate = fmt_(checkOutTime instanceof Date ? checkOutTime : new Date(checkOutTime), "MM/dd/yyyy");
      }
    }
    
    const checkIn = C.checkIn ? toYes_(data[i][C.checkIn - 1]) : false;
    const checkOut = C.checkOut ? toYes_(data[i][C.checkOut - 1]) : false;
    const hkStatus = C.hkStatus ? (data[i][C.hkStatus - 1] || "").toString() : "";
    const hkDone = C.hkDone ? toYes_(data[i][C.hkDone - 1]) : false;
    const paymentType = C.paymentType ? (data[i][C.paymentType - 1] || "").toString().toLowerCase() : "";
    const total = C.total ? toNumber_(data[i][C.total - 1]) : 0;
    
    // Check if this row is for today
    if (rowDate === todayStr || (!rowDate && (checkIn || checkOut))) {
      // Check-in today
      if (checkIn) {
        const checkInTime = C.checkInTime ? data[i][C.checkInTime - 1] : null;
        if (checkInTime) {
          const checkInDate = fmt_(checkInTime instanceof Date ? checkInTime : new Date(checkInTime), "MM/dd/yyyy");
          if (checkInDate === todayStr) {
            checkedInToday++;
          }
        } else if (rowDate === todayStr) {
          checkedInToday++;
        }
      }
      
      // Check-out today
      if (checkOut) {
        const checkOutTime = C.checkOutTime ? data[i][C.checkOutTime - 1] : null;
        if (checkOutTime) {
          const checkOutDate = fmt_(checkOutTime instanceof Date ? checkOutTime : new Date(checkOutTime), "MM/dd/yyyy");
          if (checkOutDate === todayStr) {
            checkedOutToday++;
            
            // Add to revenue
            totalRevenue += total;
            
            // Categorize by payment type
            if (paymentType.includes("cash")) {
              cashTotal += total;
            } else if (paymentType.includes("card") || paymentType.includes("credit") || paymentType.includes("debit")) {
              cardTotal += total;
            }
          }
        } else if (rowDate === todayStr) {
          checkedOutToday++;
          totalRevenue += total;
          if (paymentType.includes("cash")) {
            cashTotal += total;
          } else if (paymentType.includes("card") || paymentType.includes("credit") || paymentType.includes("debit")) {
            cardTotal += total;
          }
        }
      }
      
      // Ready for cleaning today (checked out but not cleaned)
      if (checkOut && !hkDone && hkStatus === CFG.HK_READY_TEXT) {
        const checkOutTime = C.checkOutTime ? data[i][C.checkOutTime - 1] : null;
        if (checkOutTime) {
          const checkOutDate = fmt_(checkOutTime instanceof Date ? checkOutTime : new Date(checkOutTime), "MM/dd/yyyy");
          if (checkOutDate === todayStr) {
            readyForCleaningToday++;
          }
        } else if (rowDate === todayStr) {
          readyForCleaningToday++;
        }
      }
      
      // Cleaned today
      if (hkDone) {
        const cleanedTime = C.cleanedTime ? data[i][C.cleanedTime - 1] : null;
        if (cleanedTime) {
          const cleanedDate = fmt_(cleanedTime instanceof Date ? cleanedTime : new Date(cleanedTime), "MM/dd/yyyy");
          if (cleanedDate === todayStr) {
            cleanedToday++;
          }
        } else if (rowDate === todayStr) {
          cleanedToday++;
        }
      }
    }
  }
  
  // Update dashboard values
  dashboardSheet.getRange("B5").setValue(checkedInToday);
  dashboardSheet.getRange("B6").setValue(checkedOutToday);
  dashboardSheet.getRange("B7").setValue(readyForCleaningToday);
  dashboardSheet.getRange("B8").setValue(cleanedToday);
  dashboardSheet.getRange("B9").setValue(cashTotal.toFixed(2));
  dashboardSheet.getRange("B10").setValue(cardTotal.toFixed(2));
  dashboardSheet.getRange("B11").setValue(totalRevenue.toFixed(2));
  
  // Update Cash Reconciliation - Expected Cash (from log)
  dashboardSheet.getRange("B15").setValue(cashTotal.toFixed(2));
  
  // Calculate Over/Short if Actual Cash Counted is entered
  const actualCash = dashboardSheet.getRange("B16").getValue();
  if (actualCash && typeof actualCash === "number") {
    const overShort = actualCash - cashTotal;
    dashboardSheet.getRange("B17").setValue(overShort.toFixed(2));
    
    // Color code: green if over, red if short
    if (overShort >= 0) {
      dashboardSheet.getRange("B17").setBackground("#C6EFCE"); // Green
    } else {
      dashboardSheet.getRange("B17").setBackground("#FFC7CE"); // Red
    }
  }
  
  log_("Daily Dashboard updated", { 
    date: todayStr, 
    checkedIn: checkedInToday, 
    checkedOut: checkedOutToday,
    revenue: totalRevenue 
  });
}

/* ===================== UNIFIED ROOM DASHBOARD ===================== */
function menuUnifiedRoomDashboard() {
  try {
    const dashboard = generateUnifiedRoomDashboard_();
    showUnifiedRoomDashboard_(dashboard);
  } catch (error) {
    log_("Error in menuUnifiedRoomDashboard", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error generating dashboard: " + error.toString());
  }
}

/**
 * Get room statuses from Rooms_Master sheet
 * Returns a map of room number -> status (e.g., "Maintenance", "Out of Order", "Construction", "Repair")
 * @param {Spreadsheet} ss - The spreadsheet object
 * @returns {Object} - Map of room numbers to their status from Rooms_Master
 */
function getRoomStatusesFromMaster_(ss) {
  const roomStatuses = {};
  
  try {
    const roomsSheet = ss.getSheetByName("Rooms_Master");
    if (roomsSheet) {
      const data = roomsSheet.getDataRange().getValues();
      const headers = data[0] || [];
      
      // Find Room # column and Status/Type column (flexible - checks both)
      const roomCol = headers.findIndex(h => (h || "").toString().toLowerCase().includes("room"));
      const statusCol = headers.findIndex(h => {
        const hLower = (h || "").toString().toLowerCase();
        return hLower.includes("status") || hLower.includes("type");
      });
      
      if (roomCol >= 0 && statusCol >= 0) {
        for (let i = 1; i < data.length; i++) {
          const room = (data[i][roomCol] || "").toString().trim();
          const statusRaw = (data[i][statusCol] || "").toString().trim();
          
          if (room && statusRaw) {
            // Normalize status - capitalize first letter of each word
            const statusLower = statusRaw.toLowerCase();
            let normalizedStatus = "";
            
            // Map to standard status names
            if (statusLower.includes("maintenance")) {
              normalizedStatus = "Maintenance";
            } else if (statusLower.includes("construction")) {
              normalizedStatus = "Construction";
            } else if (statusLower.includes("out of order")) {
              normalizedStatus = "Out of Order";
            } else if (statusLower.includes("repair")) {
              normalizedStatus = "Repair";
            } else if (statusLower.includes("available")) {
              normalizedStatus = "Available";
            } else {
              // Use original status with proper capitalization
              normalizedStatus = statusRaw.split(' ').map(word => 
                word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()
              ).join(' ');
            }
            
            roomStatuses[room] = normalizedStatus;
          }
        }
      }
    }
  } catch (error) {
    log_("Error reading room statuses from Rooms_Master", { error: error.toString() });
  }
  
  return roomStatuses;
}

/**
 * @deprecated Use getRoomStatusesFromMaster_ instead
 * Get maintenance/construction rooms from Rooms_Master sheet (for backward compatibility)
 */
function getMaintenanceRooms_(ss) {
  const roomStatuses = getRoomStatusesFromMaster_(ss);
  // Return only rooms that are not "Available"
  return Object.keys(roomStatuses).filter(room => roomStatuses[room] !== "Available");
}

function getRoomLabels_(ss) {
  // Get room labels from Rooms_Master sheet
  const roomLabels = {};
  
  try {
    const roomsSheet = ss.getSheetByName("Rooms_Master");
    if (roomsSheet) {
      const data = roomsSheet.getDataRange().getValues();
      const headers = data[0] || [];
      
      // Find Room # and Label columns
      const roomCol = headers.findIndex(h => (h || "").toString().toLowerCase().includes("room"));
      const labelCol = headers.findIndex(h => {
        const hLower = (h || "").toString().toLowerCase();
        return hLower.includes("label") || hLower.includes("bed") || hLower.includes("description");
      });
      
      if (roomCol >= 0 && labelCol >= 0) {
        for (let i = 1; i < data.length; i++) {
          const room = (data[i][roomCol] || "").toString().trim();
          const label = (data[i][labelCol] || "").toString().trim();
          
          if (room && label) {
            roomLabels[room] = label;
          }
        }
      }
    }
  } catch (error) {
    log_("Error reading room labels", { error: error.toString() });
  }
  
  return roomLabels;
}

function generateUnifiedRoomDashboard_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CFG.SHEET);
  
  if (!logSheet) {
    throw new Error(`Sheet not found: ${CFG.SHEET}`);
  }
  
  const C = cols_(logSheet);
  const data = logSheet.getDataRange().getValues();
  
  // Initialize room status map
  const roomStatus = {};
  const roomDetails = {};
  
  // Process all rows to get current status
  for (let i = 1; i < data.length; i++) {
    const room = C.room ? data[i][C.room - 1] : "";
    if (!room) continue;
    
    const checkIn = C.checkIn ? toYes_(data[i][C.checkIn - 1]) : false;
    const checkOut = C.checkOut ? toYes_(data[i][C.checkOut - 1]) : false;
    const hkDone = C.hkDone ? toYes_(data[i][C.hkDone - 1]) : false;
    const hkStatus = C.hkStatus ? (data[i][C.hkStatus - 1] || "").toString() : "";
    const guest = C.guest ? data[i][C.guest - 1] : "";
    const checkInTime = C.checkInTime ? data[i][C.checkInTime - 1] : "";
    const checkOutTime = C.checkOutTime ? data[i][C.checkOutTime - 1] : "";
    const nights = C.nights ? toNumber_(data[i][C.nights - 1]) : 0;
    
    // Determine status (most recent booking takes precedence)
    let status = "Available";
    let details = { guest: "", checkInTime: "", checkOutTime: "", expectedCheckOut: "" };
    
    if (checkIn && !checkOut) {
      // Calculate expected checkout time (always 11:00 AM)
      let expectedCheckOut = "";
      if (checkInTime && nights > 0) {
        try {
          const checkInDate = checkInTime instanceof Date ? checkInTime : new Date(checkInTime);
          const checkoutDate = new Date(checkInDate);
          checkoutDate.setDate(checkoutDate.getDate() + nights);
          // Set checkout time to 11:00 AM
          checkoutDate.setHours(11, 0, 0, 0);
          expectedCheckOut = fmt_(checkoutDate, "MM/dd hh:mm a");
        } catch (error) {
          log_("Error calculating checkout date", { error: error.toString() });
        }
      }
      
      status = "Occupied";
      details = {
        guest: guest || "",
        checkInTime: checkInTime ? fmt_(checkInTime, "MM/dd hh:mm a") : "",
        checkOutTime: "",
        expectedCheckOut: expectedCheckOut
      };
    } else if (checkOut && !hkDone) {
      status = "Ready for Cleaning";
      details = {
        guest: guest || "",
        checkInTime: checkInTime ? fmt_(checkInTime, "MM/dd hh:mm a") : "",
        checkOutTime: checkOutTime ? fmt_(checkOutTime, "MM/dd hh:mm a") : "",
        expectedCheckOut: ""
      };
    } else if (checkOut && hkDone) {
      // When HK Done = Yes, room is available (cleaned and ready for rent)
      status = "Available";
      details = {
        guest: "",
        checkInTime: "",
        checkOutTime: "",
        expectedCheckOut: ""
      };
    }
    
    // Keep most recent status (later rows override earlier ones)
    if (!roomStatus[room] || checkIn || checkOut) {
      roomStatus[room] = status;
      roomDetails[room] = details;
    }
  }
  
  // Always generate all 60 rooms (101-160)
  const allRooms = [];
  const roomStatusesFromMaster = getRoomStatusesFromMaster_(ss);
  const roomLabels = getRoomLabels_(ss);
  
  // Generate rooms 101-160 (60 rooms total)
  for (let r = 101; r <= 160; r++) {
    const roomStr = r.toString();
    const label = roomLabels[roomStr] || "";
    
    // Check if room has a status from Rooms_Master (Maintenance, Out of Order, Construction, Repair, etc.)
    const masterStatus = roomStatusesFromMaster[roomStr];
    
    if (masterStatus && masterStatus !== "Available") {
      // Use status from Rooms_Master (Maintenance, Out of Order, Construction, Repair, etc.)
      allRooms.push({
        number: roomStr,
        label: label,
        status: masterStatus,
        details: { guest: "", checkInTime: "", checkOutTime: "", expectedCheckOut: "" }
      });
    } else {
      // Use status from log data, or default to Available
      const details = roomDetails[roomStr] || { guest: "", checkInTime: "", checkOutTime: "", expectedCheckOut: "" };
      allRooms.push({
        number: roomStr,
        label: label,
        status: roomStatus[roomStr] || "Available",
        details: details
      });
    }
  }
  
  // Count by status - dynamically count all statuses found
  const counts = {
    Available: 0,
    Occupied: 0,
    "Ready for Cleaning": 0,
    "Cleaned - Ready": 0,
    Maintenance: 0,
    "Out of Order": 0,
    Construction: 0,
    Repair: 0
  };
  
  // Calculate checkout statistics
  const today = nowEST_();
  const todayStr = fmt_(today, "MM/dd/yyyy");
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayStr = fmt_(yesterday, "MM/dd/yyyy");
  
  let checkoutsToday = 0;
  let checkoutsYesterday = 0;
  
  // Forecast checkouts today based on check-in date + number of nights
  // This includes all bookings, not just currently occupied
  for (let i = 1; i < data.length; i++) {
    const checkIn = C.checkIn ? toYes_(data[i][C.checkIn - 1]) : false;
    const checkOut = C.checkOut ? toYes_(data[i][C.checkOut - 1]) : false;
    const checkInTime = C.checkInTime ? data[i][C.checkInTime - 1] : "";
    const nights = C.nights ? toNumber_(data[i][C.nights - 1]) : 0;
    
    // Forecast: if checked in and not checked out yet, calculate expected checkout
    if (checkIn && !checkOut && checkInTime && nights > 0) {
      try {
        const checkInDate = checkInTime instanceof Date ? checkInTime : new Date(checkInTime);
        const expectedCheckOutDate = new Date(checkInDate);
        expectedCheckOutDate.setDate(expectedCheckOutDate.getDate() + nights);
        expectedCheckOutDate.setHours(11, 0, 0, 0); // Set to 11 AM checkout time
        const expectedCheckOutStr = fmt_(expectedCheckOutDate, "MM/dd/yyyy");
        
        if (expectedCheckOutStr === todayStr) {
          checkoutsToday++;
        }
      } catch (error) {
        log_("Error calculating forecast checkout", { error: error.toString() });
      }
    }
  }
  
  // Count actual checkouts yesterday from log data
  for (let i = 1; i < data.length; i++) {
    const checkOut = C.checkOut ? toYes_(data[i][C.checkOut - 1]) : false;
    const checkOutTime = C.checkOutTime ? data[i][C.checkOutTime - 1] : "";
    
    if (checkOut && checkOutTime) {
      try {
        const checkOutDate = checkOutTime instanceof Date ? checkOutTime : new Date(checkOutTime);
        const checkOutDateStr = fmt_(checkOutDate, "MM/dd/yyyy");
        if (checkOutDateStr === yesterdayStr) {
          checkoutsYesterday++;
        }
      } catch (error) {
        log_("Error parsing checkout date", { error: error.toString() });
      }
    }
  }
  
  // Count by status
  allRooms.forEach(room => {
    counts[room.status] = (counts[room.status] || 0) + 1;
  });
  
  return {
    rooms: allRooms,
    counts: counts,
    totalRooms: allRooms.length,
    checkoutsToday: checkoutsToday,
    checkoutsYesterday: checkoutsYesterday
  };
}

function showUnifiedRoomDashboard_(dashboard) {
  const today = fmt_(nowEST_(), "MM/dd/yyyy hh:mm a");
  
  let html = `<html><head>
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; background-color: #f5f5f5; }
    .header { background-color: #2c3e50; color: white; padding: 15px; border-radius: 5px; margin-bottom: 20px; position: relative; }
    .print-btn { position: absolute; top: 15px; right: 15px; background-color: white; color: #2c3e50; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; font-weight: bold; }
    .print-btn:hover { background-color: #f0f0f0; }
    .stats { display: flex; gap: 15px; margin-bottom: 20px; flex-wrap: wrap; }
    .stat-box { background: white; padding: 15px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); min-width: 150px; }
    .stat-box h3 { margin: 0 0 10px 0; font-size: 14px; color: #666; }
    .stat-box .number { font-size: 32px; font-weight: bold; }
    .room-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(120px, 1fr)); gap: 8px; }
    .room-card { background: white; padding: 10px; border-radius: 5px; border-left: 4px solid #ddd; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .room-card.occupied { border-left-color: #FFC7CE; background-color: #FFE6E6; }
    .room-card.ready { border-left-color: #FFD700; background-color: #FFEB3B; }
    .room-card.ready-for-cleaning { border-left-color: #FFD700; background-color: #FFEB3B; }
    .room-card.cleaned { border-left-color: #D5E8D4; background-color: #E8F5E9; }
    .room-card.cleaned---ready { border-left-color: #D5E8D4; background-color: #E8F5E9; }
    .room-card.available { border-left-color: #C6EFCE; background-color: #E8F5E9; }
    .room-card.maintenance { border-left-color: #D9D9D9; background-color: #F0F0F0; }
    .room-card.out-of-order { border-left-color: #FF6B6B; background-color: #FFE0E0; }
    .room-card.construction { border-left-color: #FFA500; background-color: #FFF4E0; }
    .room-card.repair { border-left-color: #FF9800; background-color: #FFF3E0; }
    .room-number { font-weight: bold; font-size: 16px; margin-bottom: 5px; }
    .room-status { font-size: 12px; color: #666; margin-bottom: 3px; }
    .room-guest { font-size: 11px; color: #999; }
    .legend { background: white; padding: 15px; border-radius: 5px; margin-top: 20px; }
    .legend-item { display: inline-block; margin-right: 20px; }
    .legend-color { display: inline-block; width: 20px; height: 20px; border-radius: 3px; vertical-align: middle; margin-right: 5px; }
    
    @media print {
      @page {
        size: landscape;
        margin: 0.5cm;
      }
      body { 
        background-color: white; 
        padding: 0;
        margin: 0;
        font-size: 10px;
      }
      .print-btn { display: none; }
      .header { 
        page-break-after: avoid; 
        margin-bottom: 10px;
        padding: 10px;
        background-color: #2c3e50 !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
      .header h1 { font-size: 18px; margin: 0; }
      .header p { font-size: 10px; margin: 5px 0 0 0; }
      .stats { 
        page-break-inside: avoid; 
        margin-bottom: 10px;
        gap: 8px;
      }
      .stat-box { 
        padding: 8px;
        min-width: 100px;
        box-shadow: none;
        border: 1px solid #ddd;
      }
      .stat-box h3 { font-size: 10px; margin: 0 0 5px 0; }
      .stat-box .number { font-size: 24px; }
      h2 { 
        font-size: 14px; 
        margin: 10px 0 8px 0;
        page-break-after: avoid;
      }
      .room-grid { 
        display: grid;
        grid-template-columns: repeat(9, 1fr);
        gap: 4px;
        page-break-inside: avoid;
        margin-bottom: 10px;
      }
      .room-card { 
        page-break-inside: avoid; 
        break-inside: avoid;
        padding: 6px;
        border-radius: 2px;
        box-shadow: none;
        border: 1px solid #ddd;
        min-height: auto;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
      .room-number { 
        font-size: 11px; 
        margin-bottom: 3px;
        font-weight: bold;
      }
      .room-status { 
        font-size: 9px; 
        margin-bottom: 2px;
      }
      .room-guest { 
        font-size: 8px;
        line-height: 1.2;
      }
      .legend { 
        page-break-before: avoid;
        padding: 8px;
        margin-top: 10px;
        font-size: 9px;
        box-shadow: none;
        border: 1px solid #ddd;
      }
      .legend h3 { font-size: 10px; margin: 0 0 5px 0; }
      .legend-item { font-size: 8px; margin-right: 15px; }
      .legend-color { width: 12px; height: 12px; }
    }
  </style>
  <script>
    function printDashboard() {
      window.print();
    }
  </script>
  </head><body>`;
  
  html += `<div class="header">
    <h1 style="margin:0;">Unified Room Dashboard</h1>
    <p style="margin:5px 0 0 0;">Total Rooms: ${dashboard.totalRooms} | Generated: ${today}</p>
    <button class="print-btn" onclick="printDashboard()">🖨️ Print</button>
  </div>`;
  
  // Statistics - show all status types
  html += `<div class="stats">`;
  html += `<div class="stat-box"><h3>Available</h3><div class="number" style="color:#27ae60;">${dashboard.counts.Available || 0}</div></div>`;
  html += `<div class="stat-box"><h3>Occupied</h3><div class="number" style="color:#e74c3c;">${dashboard.counts.Occupied || 0}</div></div>`;
  html += `<div class="stat-box"><h3>Ready for Cleaning</h3><div class="number" style="color:#f39c12;">${dashboard.counts["Ready for Cleaning"] || 0}</div></div>`;
  html += `<div class="stat-box"><h3>Cleaned - Ready</h3><div class="number" style="color:#27ae60;">${dashboard.counts["Cleaned - Ready"] || 0}</div></div>`;
  
  // Show maintenance-related statuses if they have counts
  if (dashboard.counts.Maintenance > 0) {
    html += `<div class="stat-box"><h3>Maintenance</h3><div class="number" style="color:#95a5a6;">${dashboard.counts.Maintenance}</div></div>`;
  }
  if (dashboard.counts["Out of Order"] > 0) {
    html += `<div class="stat-box"><h3>Out of Order</h3><div class="number" style="color:#e74c3c;">${dashboard.counts["Out of Order"]}</div></div>`;
  }
  if (dashboard.counts.Construction > 0) {
    html += `<div class="stat-box"><h3>Construction</h3><div class="number" style="color:#f39c12;">${dashboard.counts.Construction}</div></div>`;
  }
  if (dashboard.counts.Repair > 0) {
    html += `<div class="stat-box"><h3>Repair</h3><div class="number" style="color:#f39c12;">${dashboard.counts.Repair}</div></div>`;
  }
  
  html += `<div class="stat-box"><h3>Checkouts Today</h3><div class="number" style="color:#3498db;">${dashboard.checkoutsToday || 0}</div></div>`;
  html += `<div class="stat-box"><h3>Checked Out Yesterday</h3><div class="number" style="color:#9b59b6;">${dashboard.checkoutsYesterday || 0}</div></div>`;
  html += `</div>`;
  
  // Room grid
  html += `<h2>Room Status</h2><div class="room-grid">`;
  
  dashboard.rooms.forEach(room => {
    const statusClass = room.status.toLowerCase().replace(/\s+/g, "-").replace(/[^a-z-]/g, "");
    html += `<div class="room-card ${statusClass}">`;
    
    // Display room number with label if available
    let roomDisplay = `Room ${room.number}`;
    if (room.label) {
      roomDisplay = `RM${room.number} ${room.label}`;
    }
    
    html += `<div class="room-number">${roomDisplay}</div>`;
    html += `<div class="room-status">${room.status}</div>`;
    if (room.details.expectedCheckOut && room.status === "Occupied") {
      html += `<div class="room-guest" style="font-size:10px; color:#e74c3c; font-weight:bold;">Out: ${room.details.expectedCheckOut}</div>`;
    }
    html += `</div>`;
  });
  
  html += `</div>`;
  
  // Legend - include all status types
  html += `<div class="legend"><h3>Legend:</h3>`;
  html += `<span class="legend-item"><span class="legend-color" style="background-color:#C6EFCE;"></span>Available</span>`;
  html += `<span class="legend-item"><span class="legend-color" style="background-color:#FFC7CE;"></span>Occupied</span>`;
  html += `<span class="legend-item"><span class="legend-color" style="background-color:#FFEB3B;"></span>Ready for Cleaning</span>`;
  html += `<span class="legend-item"><span class="legend-color" style="background-color:#D5E8D4;"></span>Cleaned - Ready</span>`;
  html += `<span class="legend-item"><span class="legend-color" style="background-color:#D9D9D9;"></span>Maintenance</span>`;
  html += `<span class="legend-item"><span class="legend-color" style="background-color:#FFE0E0;"></span>Out of Order</span>`;
  html += `<span class="legend-item"><span class="legend-color" style="background-color:#FFF4E0;"></span>Construction</span>`;
  html += `<span class="legend-item"><span class="legend-color" style="background-color:#FFF3E0;"></span>Repair</span>`;
  html += `</div>`;
  
  html += `</body></html>`;
  
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Unified Room Dashboard");
}

/* ===================== ROOM MAINTENANCE MANAGEMENT ===================== */
function menuManageRoomMaintenance() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let roomsSheet = ss.getSheetByName("Rooms_Master");
    
    // Create Rooms_Master sheet if it doesn't exist
    if (!roomsSheet) {
      roomsSheet = ss.insertSheet("Rooms_Master");
      roomsSheet.getRange(1, 1).setValue("Room #");
      roomsSheet.getRange(1, 2).setValue("Type");
      roomsSheet.getRange(1, 3).setValue("Notes");
      roomsSheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
      
      // Add all 60 rooms (101-160) with default status
      const rows = [];
      for (let r = 101; r <= 160; r++) {
        rows.push([r.toString(), "Available", ""]);
      }
      if (rows.length > 0) {
        roomsSheet.getRange(2, 1, rows.length, 3).setValues(rows);
      }
      
      // Add data validation for Type column
      const typeRange = roomsSheet.getRange(2, 2, rows.length, 1);
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Available", "Maintenance", "Construction", "Out of Order", "Repair"], true)
        .build();
      typeRange.setDataValidation(rule);
      
      SpreadsheetApp.getUi().alert("Rooms_Master sheet created! Please mark rooms as 'Maintenance' or 'Construction' in column B (Type).");
    } else {
      // Sheet exists - check if it needs data validation
      const headers = roomsSheet.getRange(1, 1, 1, roomsSheet.getLastColumn()).getValues()[0];
      const typeCol = headers.findIndex(h => {
        const hLower = (h || "").toString().toLowerCase();
        return hLower.includes("status") || hLower.includes("type");
      });
      
      if (typeCol >= 0) {
        const lastRow = roomsSheet.getLastRow();
        if (lastRow > 1) {
          const typeRange = roomsSheet.getRange(2, typeCol + 1, lastRow - 1, 1);
          // Check if validation already exists
          const existingRule = typeRange.getDataValidation();
          if (!existingRule) {
            const rule = SpreadsheetApp.newDataValidation()
              .requireValueInList(["Available", "Maintenance", "Construction", "Out of Order", "Repair"], true)
              .build();
            typeRange.setDataValidation(rule);
          }
        }
      }
    }
    
    // Activate the sheet
    ss.setActiveSheet(roomsSheet);
    SpreadsheetApp.getUi().alert("Rooms_Master sheet opened. Update room types in column B:\n- Maintenance\n- Construction\n- Out of Order\n- Repair\n\nAvailable rooms should be left as 'Available'.");
    
  } catch (error) {
    log_("Error in menuManageRoomMaintenance", { error: error.toString() });
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

/* ===================== PENDING HOUSEKEEPING REPORT ===================== */
function showPendingHousekeepingReport_(items) {
  const today = fmt_(nowEST_(), "MM/dd/yyyy hh:mm a");
  
  let html = `<html><head>
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; background-color: #ffffff; }
    .header { text-align: center; margin-bottom: 25px; border-bottom: 2px solid #333; padding-bottom: 15px; }
    .motel-name { font-size: 24px; font-weight: bold; color: #2c3e50; margin: 0; }
    .motel-address { margin-top: 8px; color: #555; font-size: 13px; }
    .motel-contact { margin-top: 8px; color: #666; font-size: 12px; }
    .report-title { font-size: 20px; font-weight: bold; margin: 20px 0 10px 0; color: #2c3e50; }
    .report-date { font-size: 12px; color: #666; margin-bottom: 20px; }
    .print-btn { position: fixed; top: 20px; right: 20px; background-color: #2c3e50; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-weight: bold; z-index: 1000; }
    .print-btn:hover { background-color: #34495e; }
    table { width: 100%; border-collapse: collapse; margin-top: 20px; }
    th { background-color: #2c3e50; color: white; padding: 12px; text-align: left; font-weight: bold; border: 1px solid #1a252f; }
    td { padding: 10px; border: 1px solid #ddd; }
    tr:nth-child(even) { background-color: #f9f9f9; }
    tr:hover { background-color: #f5f5f5; }
    .date-column { background-color: #fffacd; min-width: 120px; }
    .no-data { text-align: center; padding: 40px; color: #999; font-style: italic; }
    
    @media print {
      body { padding: 10px; }
      .print-btn { display: none; }
      .header { page-break-after: avoid; }
      table { page-break-inside: avoid; }
      tr { page-break-inside: avoid; }
    }
  </style>
  <script>
    function printReport() {
      window.print();
    }
  </script>
  </head><body>`;
  
  // Motel header
  html += `<div class="header">
    <h1 class="motel-name">${CFG.MOTEL.name}</h1>
    <div class="motel-address">${CFG.MOTEL.addr1}<br>${CFG.MOTEL.addr2}</div>
    <div class="motel-contact">Phone: ${CFG.MOTEL.phone} | Email: ${CFG.MOTEL.email}</div>
  </div>`;
  
  html += `<button class="print-btn" onclick="printReport()">🖨️ Print</button>`;
  
  html += `<div class="report-title">Pending Housekeeping</div>`;
  html += `<div class="report-date">Generated: ${today}</div>`;
  
  if (items.length === 0) {
    html += `<div class="no-data">No pending housekeeping tasks at this time.</div>`;
  } else {
    html += `<table>
      <thead>
        <tr>
          <th>Room</th>
          <th>Guest</th>
          <th>Check Out Time</th>
          <th class="date-column">Date Cleaned</th>
        </tr>
      </thead>
      <tbody>`;
    
    for (const item of items) {
      html += `<tr>`;
      html += `<td><strong>${item.room || "-"}</strong></td>`;
      html += `<td>${item.guest || "-"}</td>`;
      html += `<td>${item.checkOutTime || "-"}</td>`;
      html += `<td class="date-column">&nbsp;</td>`; // Blank column for housekeeping to fill
      html += `</tr>`;
    }
    
    html += `</tbody></table>`;
    
    html += `<div style="margin-top: 20px; font-size: 11px; color: #666; text-align: center;">
      Total Rooms Pending: ${items.length}
    </div>`;
  }
  
  html += `</body></html>`;
  
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(900).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Pending Housekeeping Report");
}

/* ===================== LOGGING ===================== */
function log_(message, data) {
  if (!CFG.LOG_ENABLED) return;
  try {
    const logData = {
      timestamp: fmt_(nowEST_(), "yyyy-MM-dd HH:mm:ss"),
      message: message,
      data: data || {}
    };
    Logger.log(JSON.stringify(logData));
  } catch (error) {
    // Silently fail logging
  }
}
