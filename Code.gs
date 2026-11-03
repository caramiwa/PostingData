/**
 * Import Purchase Orders
 *
 * Imports only:
 *  - STATUS = "Purchase Order"
 *  - DATE PREPARED (PO) >= MIN_DATE
 *
 * Consolidates all source rows having the same PO/CONTRACT NO.
 * Existing PO/CONTRACT NO. records are NOT overwritten.
 * Blank cells in existing records are highlighted.
 *
 * SOURCE SHEET IS READ ONLY.
 *
 * For each PO/CONTRACT NO.:
 *  - PROJECT ID is normalized to the value inside parentheses.
 *  - Text/date fields keep distinct values, separated by line breaks.
 *  - PO TOTAL COST is summed across all qualifying source rows.
 */
function importPurchaseOrders() {

  const SOURCE_ID = "1xb76Hf1uV600UyCIuev1iBupolLWR9V2KVx9FffK3SU";
  const SOURCE_SHEET = "Data List";
  const DEST = SpreadsheetApp.getActiveSheet();

  const MIN_DATE = new Date(2026, 6, 1); // July 1, 2026

  const DEST_HEADERS = [
    "PROJECT ID",
    "PROJECT TITLE",
    "BAC RESOLUTION NO.",
    "NOA ISSUANCE DATE",
    "NTP ISSUANCE DATE",
    "PO/CONTRACT NO.",
    "TOTAL CONTRACT AMOUNT",
    "BIDDERS/SUPPLIERS"
  ];

  const FIELD_MAP = {
    "PROJECT ID": "PROJECT ID",
    "PROJECT TITLE": "PROJECT TITLE",
    "BAC RESOLUTION NO.": "BAC RESOLUTION NO.",
    "NOA DATE": "NOA ISSUANCE DATE",
    "NTP DATE": "NTP ISSUANCE DATE",
    "PO NO.": "PO/CONTRACT NO.",
    "PO TOTAL COST": "TOTAL CONTRACT AMOUNT",
    "SUPPLIER": "BIDDERS/SUPPLIERS"
  };

  const DATE_FIELDS = ["NOA DATE", "NTP DATE"];

  // ------------------------------------------------------------
  // SOURCE: READ ONLY
  // ------------------------------------------------------------

  const srcSheet = SpreadsheetApp.openById(SOURCE_ID).getSheetByName(SOURCE_SHEET);
  if (!srcSheet) {
    throw new Error(`Source sheet "${SOURCE_SHEET}" was not found.`);
  }

  const srcValues = srcSheet.getDataRange().getValues();
  if (srcValues.length < 2) return;

  const srcHeaders = srcValues[0];
  const srcData = srcValues.slice(1);

  const srcCol = {};
  srcHeaders.forEach((h, i) => {
    srcCol[String(h).trim()] = i;
  });

  [
    "PROJECT ID",
    "PROJECT TITLE",
    "BAC RESOLUTION NO.",
    "NOA DATE",
    "NTP DATE",
    "PO NO.",
    "PO TOTAL COST",
    "SUPPLIER",
    "STATUS",
    "DATE PREPARED (PO)"
  ].forEach(header => {
    if (srcCol[header] === undefined) {
      throw new Error(`Source column "${header}" was not found.`);
    }
  });

  // ------------------------------------------------------------
  // Consolidate qualifying source rows by PO/CONTRACT NO.
  // ------------------------------------------------------------

  const grouped = {};

  srcData.forEach(row => {

    if (String(row[srcCol["STATUS"]]).trim() !== "Purchase Order") {
      return;
    }

    const poDate = parseFlexibleDate(row[srcCol["DATE PREPARED (PO)"]]);

    if (!poDate || poDate < MIN_DATE) {
      return;
    }

    const poNo = normalizeValue(row[srcCol["PO NO."]]);
    if (!poNo) return;

    if (!grouped[poNo]) {
      grouped[poNo] = {
        "PROJECT ID": [],
        "PROJECT TITLE": [],
        "BAC RESOLUTION NO.": [],
        "NOA DATE": [],
        "NTP DATE": [],
        "PO NO.": poNo,
        "PO TOTAL COST": 0,
        "SUPPLIER": []
      };
    }

    // Normalize text/date fields. The source is never written to.
    [
      "PROJECT ID",
      "PROJECT TITLE",
      "BAC RESOLUTION NO.",
      "NOA DATE",
      "NTP DATE",
      "SUPPLIER"
    ].forEach(field => {

      let value = row[srcCol[field]];

      if (field === "PROJECT ID") {
        value = extractProjectID(value);
      }

      if (DATE_FIELDS.includes(field)) {
        value = formatDisplayDate(value);
      }

      value = normalizeValue(value);

      if (value && !grouped[poNo][field].includes(value)) {
        grouped[poNo][field].push(value);
      }
    });

    // SUM all PO TOTAL COST values belonging to this PO/CONTRACT NO.
    grouped[poNo]["PO TOTAL COST"] += parseAmountTotal(
      row[srcCol["PO TOTAL COST"]]
    );
  });

  // ------------------------------------------------------------
  // DESTINATION
  // ------------------------------------------------------------

  if (DEST.getLastRow() === 0) {
    DEST.getRange(1, 1, 1, DEST_HEADERS.length).setValues([DEST_HEADERS]);
  }

  const destData = DEST.getDataRange().getValues();
  const destHeaders = destData[0];

  const destCol = {};
  destHeaders.forEach((h, i) => {
    destCol[String(h).trim()] = i;
  });

  if (destCol["PO/CONTRACT NO."] === undefined) {
    throw new Error("Destination column \"PO/CONTRACT NO.\" was not found.");
  }

  // ------------------------------------------------------------
  // Build target index by PO/CONTRACT NO.
  // ------------------------------------------------------------

  const existing = {};

  for (let r = 1; r < destData.length; r++) {

    const poNo = normalizeValue(destData[r][destCol["PO/CONTRACT NO."]]);

    if (poNo) {
      existing[poNo] = r + 1;
    }
  }

  let added = 0;
  let skipped = 0;
  let incomplete = 0;

  // ------------------------------------------------------------
  // Write one row per PO/CONTRACT NO.
  // ------------------------------------------------------------

  Object.keys(grouped).forEach(poNo => {

    const group = grouped[poNo];

    // Existing PO/Contract: never overwrite it.
    if (existing[poNo]) {

      skipped++;

      const sheetRow = existing[poNo];
      let hasBlank = false;

      DEST_HEADERS.forEach(header => {

        const col = destCol[header];
        if (col === undefined) return;

        const cell = DEST.getRange(sheetRow, col + 1);

        if (normalizeValue(cell.getValue()) === "") {
          cell.setBackground("#FFF2CC");
          hasBlank = true;
        } else {
          cell.setBackground(null);
        }
      });

      if (hasBlank) incomplete++;
      return;
    }

    // Build the target row explicitly so the consolidation key can
    // never be omitted by the generic field mapping.
    const newRow = [
      group["PROJECT ID"].join("\n"),
      group["PROJECT TITLE"].join("\n"),
      group["BAC RESOLUTION NO."].join("\n"),
      group["NOA DATE"].join("\n"),
      group["NTP DATE"].join("\n"),
      group["PO NO."],
      group["PO TOTAL COST"],
      group["SUPPLIER"].join("\n")
    ];

    const targetRow = DEST.getLastRow() + 1;
    DEST.getRange(targetRow, 1, 1, newRow.length).setValues([newRow]);

    // Display dates as: Jul 17 2026
    if (destCol["NOA ISSUANCE DATE"] !== undefined) {
      DEST.getRange(targetRow, destCol["NOA ISSUANCE DATE"] + 1)
        .setNumberFormat("mmm d yyyy");
    }

    if (destCol["NTP ISSUANCE DATE"] !== undefined) {
      DEST.getRange(targetRow, destCol["NTP ISSUANCE DATE"] + 1)
        .setNumberFormat("mmm d yyyy");
    }

    // Display summed contract amount with thousands separators and 2 decimals.
    if (destCol["TOTAL CONTRACT AMOUNT"] !== undefined) {
      DEST.getRange(targetRow, destCol["TOTAL CONTRACT AMOUNT"] + 1)
        .setNumberFormat("#,##0.00");
    }

    added++;
  });

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    `Purchase Order Import Complete\n\n` +
    `New consolidated records : ${added}\n` +
    `Existing PO/Contract records : ${skipped}\n` +
    `Incomplete existing rows : ${incomplete}`
  );
}


function normalizeValue(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}


function parseFlexibleDate(value) {

  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }

  if (!value) return null;

  value = String(value).trim();

  let m = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);

  m = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[1] - 1, +m[2]);

  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}


function formatDisplayDate(value) {

  const date = parseFlexibleDate(value);
  if (!date) return "";

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "MMM d yyyy"
  );
}


function parseAmountTotal(value) {

  if (value === null || value === undefined || value === "") {
    return 0;
  }

  // The source may contain one amount or multiple line-separated amounts.
  // Each line is treated as a component and summed.
  const parts = String(value)
    .split(/\r?\n/)
    .map(part => part.trim())
    .filter(Boolean);

  let total = 0;

  parts.forEach(part => {
    const cleaned = part.replace(/[₱,\s]/g, "");
    const amount = Number(cleaned);

    if (!isNaN(amount)) {
      total += amount;
    }
  });

  return total;
}


function extractProjectID(value) {

  if (!value) return "";

  value = String(value).trim();

  const start = value.indexOf("(");
  const end = value.lastIndexOf(")");

  if (start >= 0 && end > start) {
    return value.substring(start + 1, end).trim();
  }

  return value;
}
