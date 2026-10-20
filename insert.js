/* Step 2: turn the deduped OTC list into the Salesforce "OTC Add Only" insert file.
   Row 1 notes and row 2 API names match Cody's template exactly. */
(function (root) {
  var PREFIX = "One Time Care Package Support: ";
  var PROGRAM_ID = "a1N4x000006CkKlEAK";     // Don't Change (OTC ID)
  var RECORD_TYPE_ID = "0124x000000aJl3AAE"; // Don't Change (OTC ID)
  var SHEET = "Step 1 - Engagement Create";
  var DQ = "Don’t Change";

  var NOTES = [
    "Update to SM Ifor (One Time Care Package Support: First Name Last Initial.)",
    "Put to Start Date", DQ, null, null, "Update to Client ID",
    DQ + " (OTC ID)", DQ + " (OTC ID)", DQ, DQ, DQ, DQ, DQ
  ];
  var HEADERS = [
    "Name", "pmdm__StartDate__c", "pmdm__Stage__c", "Branch_of_Service__c", "Message_to_Angels__c",
    "pmdm__Contact__c", "pmdm__Program__c", "RecordTypeId", "Volunteer_Capacity__c",
    "Registered_Volunteers__c", "Volunteer_Capacity_Correct__c",
    "Requested_Wishlist_Quantities_Set__c", "Documents_Provided__c"
  ];
  var WIDTHS = [75.14, 27.14, 22.86, 28.86, 30.57, 25.57, 26.14, 24, 29.14, 33.14, 40.14, 49, 31.86];
  // Data Loader settings that sit beside the data on the template.
  var SIDE = [["O3", "API Type"], ["P3", "Enterprise"], ["R3", "RecordType ID"], ["O4", "Action"], ["P4", "Insert"],
              ["O5", "Object"], ["P5", "pmdm__ProgramEngagement__c"], ["O6", "Data"], ["O7", "Account"]];

  function txt(v) { return v == null ? "" : String(v).replace(/\s+/g, " ").trim(); }

  // Fix ALL CAPS or all lowercase words; leave mixed case (McDonald, DeShawn) alone.
  function fixWord(w) {
    if (!/[A-Za-z]/.test(w)) return w;
    if (w !== w.toUpperCase() && w !== w.toLowerCase()) return w;
    return w.toLowerCase().replace(/(^|[-'’])([a-z])/g, function (m, p, c) { return p + c.toUpperCase(); });
  }

  // "First Name Last Initial." Salesforce's own engagement name already splits first and last
  // correctly ("Eric Cruz De Paz" -> "Eric C."), so use it and just fix the casing.
  function shortName(engagement, fullName) {
    var s = txt(engagement);
    if (s.indexOf(":") > -1) s = txt(s.slice(s.indexOf(":") + 1));
    if (!/\s[A-Za-z]\.$/.test(s)) {
      var parts = txt(fullName).split(" ").filter(Boolean);
      if (parts.length < 2) return parts.length ? fixWord(parts[0]) : "";
      s = parts[0] + " " + parts[parts.length - 1].charAt(0) + ".";
    }
    return s.split(" ").map(function (w) {
      return /^[a-z]\.$/i.test(w) ? w.toUpperCase() : fixWord(w);
    }).join(" ");
  }

  function findCol(header, tests) {
    var h = header.map(function (x) { return txt(x).toLowerCase(); });
    for (var t = 0; t < tests.length; t++) {
      for (var i = 0; i < h.length; i++) if (tests[t].test(h[i])) return i;
    }
    return -1;
  }

  function mapColumns(header) {
    return {
      id: findCol(header, [/contact\s*id/, /\bid\b/]),
      full: findCol(header, [/full\s*name/, /client.*name/]),
      eng: findCol(header, [/program engagement name/, /engagement/]),
      branch: findCol(header, [/branch/]),
      msg: findCol(header, [/message to angels/, /message/])
    };
  }

  function buildRows(header, rows) {
    var c = mapColumns(header), out = [], warnings = [], byName = {};
    var get = function (r, i) { return i < 0 ? "" : txt(r[i]); };
    if (c.id < 0) warnings.push("Could not find the Client: Contact ID column.");
    if (c.full < 0 && c.eng < 0) warnings.push("Could not find a name column.");
    rows.forEach(function (r) {
      var id = get(r, c.id);
      if (!id) return;
      var nm = shortName(get(r, c.eng), get(r, c.full));
      var item = { short: nm, name: nm ? PREFIX + nm : "", fullName: get(r, c.full),
                   branch: get(r, c.branch), msg: get(r, c.msg), contact: id };
      if (!nm) warnings.push("No name for Contact ID " + id + ". Fill in column A before uploading.");
      var k = item.fullName.toLowerCase() || nm.toLowerCase();
      (byName[k] = byName[k] || []).push(id);
      out.push(item);
    });
    Object.keys(byName).forEach(function (k) {
      if (byName[k].length > 1) warnings.push("Same name, different Contact IDs: " +
        byName[k].join(" and ") + ". Both stayed in. Check they are two different people.");
    });
    return { items: out, warnings: warnings, columns: c };
  }

  // startDate is "YYYY-MM-DD".
  function buildWorkbook(ExcelJS, items, startDate) {
    var p = startDate.split("-").map(Number);
    var date = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
    var wb = new ExcelJS.Workbook();
    var ws = wb.addWorksheet(SHEET);
    var font = { name: "Aptos Narrow", size: 11 };
    ws.columns = WIDTHS.map(function (w) { return { width: w }; });
    ws.getRow(1).values = NOTES;
    ws.getRow(2).values = HEADERS;
    items.forEach(function (it, i) {
      ws.getRow(i + 3).values = [it.name, date, "Active", it.branch || null, it.msg || null, it.contact,
        PROGRAM_ID, RECORD_TYPE_ID, 1, 0, true, true, true];
    });
    SIDE.forEach(function (s) { ws.getCell(s[0]).value = s[1]; });
    var last = Math.max(items.length + 2, 7);
    for (var r = 1; r <= last; r++) {
      var row = ws.getRow(r);
      for (var col = 1; col <= 18; col++) row.getCell(col).font = font;
    }
    ws.getRow(2).eachCell(function (cell) {
      cell.font = { name: "Aptos Narrow", size: 11, bold: true };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
    });
    ws.getColumn(2).numFmt = "yyyy\\-mm\\-dd;@";
    return wb;
  }

  var api = { PREFIX: PREFIX, HEADERS: HEADERS, NOTES: NOTES, shortName: shortName,
              buildRows: buildRows, buildWorkbook: buildWorkbook };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.OTCInsert = api;
})(this);
