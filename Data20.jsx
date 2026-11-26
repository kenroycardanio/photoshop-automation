/*
  Photoshop CS6 — Data20.jsx — NFT back fill (client) — Data20: DataA-DataT
  Groups id1…id5. Sheet: NFT CARDS
  Col1 # = row number (ignored). Col2 = Card Number.
  PictureA→smartA1…  PictureB→smartB1…
  DataA→dataA1… DataT (text). Member CSV: nickname/firstName/lastName/verifyUrl.
  Card Number kept as 8155 5265 3090 3026 (4-4-4-4).
  Headers match case-insensitively, ignoring spaces and underscores.
  Rows field: 102-104 or 102,103,107. 1 = first data row after the header (Excel/Sheets row 2); # column ignored.
  CSV encoding: try UTF-8, then SYSTEM/Windows-1252 if invalid UTF-8.
*/

#target photoshop
app.bringToFront();

(function () {
  if (!app.documents.length) {
    alert("Open your NFT template PSD first, then run this script.");
    return;
  }

  var FIELDS = [
    {
      prefix: "cardnumber",
      column: "Card Number",
      type: "cardnumber",
      columnAliases: ["CardNumber", "CARD NUMBER", "card number"]
    },
    {
      prefix: "qrcardnumber",
      column: "QR Card Number",
      type: "digits"
    },
    {
      prefix: "name",
      column: "NAME",
      type: "text-fit-track",
      fitRules: {
        fullAt: 15,
        fullTrack: 0,
        anchorAt: 20,
        anchorTrack: 0,
        minTrack: 0,
        scaleFullAt: 15,
        scaleFull: 100,
        midAt: 20,
        midScale: 95,
        scaleAnchorAt: 25,
        scaleAnchor: 80,
        minScale: 80
      }
    },
    { prefix: "membersince", column: "MEMBER SINCE", type: "text" },
    { prefix: "validuntil", column: "Valid Until", type: "text" },
    // PictureC–E stay disabled (no layer hunt / no skip spam)
    {
      prefix: "smartA",
      column: "PictureA",
      type: "image",
      columnAliases: [
        "Picture A",
        "pictureA",
        "PICTUREA",
        "smart1path",
        "Smart1 Path",
        "Data1"
      ],
      layerAliases: ["smartA", "smartA1", "smart1", "qr"]
    },
    {
      prefix: "smartB",
      column: "PictureB",
      type: "image",
      optional: true,
      columnAliases: ["Picture B", "smart2path"],
      layerAliases: ["smartB"]
    },
    {
      prefix: "smartC",
      column: "PictureC",
      type: "image",
      enabled: false,
      columnAliases: ["Picture C", "smart3path"],
      layerAliases: ["smartC"]
    },
    {
      prefix: "smartD",
      column: "PictureD",
      type: "image",
      enabled: false,
      columnAliases: ["Picture D", "smart4path"],
      layerAliases: ["smartD"]
    },
    {
      prefix: "smartE",
      column: "PictureE",
      type: "image",
      enabled: false,
      columnAliases: ["Picture E", "smart5path"],
      layerAliases: ["smartE"]
    },
    {
      prefix: "dataA",
      column: "DataA",
      type: "text",
      optional: true,
      columnAliases: ["Data A"],
      layerAliases: ["dataA"]
    },
    {
      prefix: "dataB",
      column: "DataB",
      type: "text",
      optional: true,
      columnAliases: ["Data B"],
      layerAliases: ["dataB"]
    },
    {
      prefix: "dataC",
      column: "DataC",
      type: "text",
      optional: true,
      columnAliases: ["Data C"],
      layerAliases: ["dataC"]
    },
    {
      prefix: "dataD",
      column: "DataD",
      type: "text",
      optional: true,
      columnAliases: ["Data D"],
      layerAliases: ["dataD"]
    },
    {
      prefix: "dataE",
      column: "DataE",
      type: "text",
      optional: true,
      columnAliases: ["Data E"],
      layerAliases: ["dataE"]
    },
    {
      prefix: "dataF",
      column: "DataF",
      type: "text",
      optional: true,
      columnAliases: ["Data F"],
      layerAliases: ["dataF"]
    },
    {
      prefix: "dataG",
      column: "DataG",
      type: "text",
      optional: true,
      columnAliases: ["Data G"],
      layerAliases: ["dataG"]
    },
    {
      prefix: "dataH",
      column: "DataH",
      type: "text",
      optional: true,
      columnAliases: ["Data H"],
      layerAliases: ["dataH"]
    },
    {
      prefix: "dataI",
      column: "DataI",
      type: "text",
      optional: true,
      columnAliases: ["Data I"],
      layerAliases: ["dataI"]
    },
    {
      prefix: "dataJ",
      column: "DataJ",
      type: "text",
      optional: true,
      columnAliases: ["Data J"],
      layerAliases: ["dataJ"]
    },
    {
      prefix: "dataK",
      column: "DataK",
      type: "text",
      optional: true,
      columnAliases: ["Data K"],
      layerAliases: ["dataK"]
    },
    {
      prefix: "dataL",
      column: "DataL",
      type: "text",
      optional: true,
      columnAliases: ["Data L"],
      layerAliases: ["dataL"]
    },
    {
      prefix: "dataM",
      column: "DataM",
      type: "text",
      optional: true,
      columnAliases: ["Data M"],
      layerAliases: ["dataM"]
    },
    {
      prefix: "dataN",
      column: "DataN",
      type: "text",
      optional: true,
      columnAliases: ["Data N"],
      layerAliases: ["dataN"]
    },
    {
      prefix: "dataO",
      column: "DataO",
      type: "text",
      optional: true,
      columnAliases: ["Data O"],
      layerAliases: ["dataO"]
    },
    {
      prefix: "dataP",
      column: "DataP",
      type: "text",
      optional: true,
      columnAliases: ["Data P"],
      layerAliases: ["dataP"]
    },
    {
      prefix: "dataQ",
      column: "DataQ",
      type: "text",
      optional: true,
      columnAliases: ["Data Q"],
      layerAliases: ["dataQ"]
    },
    {
      prefix: "dataR",
      column: "DataR",
      type: "text",
      optional: true,
      columnAliases: ["Data R"],
      layerAliases: ["dataR"]
    },
    {
      prefix: "dataS",
      column: "DataS",
      type: "text",
      optional: true,
      columnAliases: ["Data S"],
      layerAliases: ["dataS"]
    },
    {
      prefix: "dataT",
      column: "DataT",
      type: "text",
      optional: true,
      columnAliases: ["Data T"],
      layerAliases: ["dataT"]
    },
    {
      prefix: "nickname",
      column: "nickname",
      type: "text",
      optional: true,
      columnAliases: ["Nick Name", "Nick"]
    },
    {
      prefix: "firstName",
      column: "firstName",
      type: "text",
      optional: true,
      columnAliases: ["First Name", "Firstname"]
    },
    {
      prefix: "lastName",
      column: "lastName",
      type: "text",
      optional: true,
      columnAliases: ["Last Name", "Lastname"]
    },
    {
      prefix: "verifyUrl",
      column: "verifyUrl",
      type: "text",
      optional: true,
      columnAliases: ["Verify URL", "verify_url", "VerifyUrl"]
    }
  ];

  var MAX_SLOTS = 5;

  function showDialog() {
    var slotCount = countIdGroups(app.activeDocument);
    var dlg = new Window("dialog", "NFT Back Fill Data20 (CS6)");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.spacing = 10;
    dlg.margins = 16;

    var fileGroup = dlg.add("group");
    fileGroup.add("statictext", undefined, "CSV:");
    var fileEt = fileGroup.add("edittext", undefined, "", { truncate: "middle" });
    fileEt.preferredSize.width = 320;
    fileGroup.add("button", undefined, "Browse…").onClick = function () {
      var f = File.openDialog(
        "Select NFT CSV (default) or Excel",
        "CSV:*.csv,Excel:*.xlsx;*.xls,All files:*.*"
      );
      if (f) fileEt.text = f.fsName;
    };

    var rowsGroup = dlg.add("group");
    rowsGroup.add("statictext", undefined, "Rows:");
    var rowsEt = rowsGroup.add("edittext", undefined, "1-5");
    rowsEt.preferredSize.width = 200;
    rowsGroup.add("statictext", undefined, "e.g. 102-104 or 102,103,107");

    dlg.add(
      "statictext",
      undefined,
      "Row numbers: 1 = first data row after the header (Excel/Sheets row 2).\n" +
        "The # column is ignored. Type a list or range in Rows (spaces OK).\n" +
        "Each listed row fills the next id group (id1, id2, …) in the order you typed.\n" +
        "This PSD has " +
        slotCount +
        " id group(s). Too many rows is rejected. Past-end or blank rows are skipped.\n" +
        "Data20: DataA-DataT. Col2 = Card Number. PictureA → smartA1… DataA–DataT fill text layers if present.\n" +
        "CSV: UTF-8, or older Windows-1252 auto-detected. Missing images are skipped and listed.",
      { multiline: true }
    ).preferredSize.width = 420;

    var btns = dlg.add("group");
    btns.alignment = "right";
    btns.add("button", undefined, "Cancel", { name: "cancel" });
    btns.add("button", undefined, "Run", { name: "ok" });

    if (dlg.show() !== 1) return null;

    var parsed = parseRowList(rowsEt.text);
    if (parsed.error) {
      alert(
        "Rows must be like 102-104 or 102,103,107 (digits, commas, and hyphens only; min 1)."
      );
      return null;
    }
    if (!fileEt.text) {
      alert("Select a CSV file (or Excel).");
      return null;
    }
    var nCards = parsed.rows.length;
    if (nCards > slotCount) {
      alert(
        "Rows lists " +
          nCards +
          " cards (" +
          parsed.rows.join(", ") +
          "), but this PSD has " +
          slotCount +
          " id group(s) (id1…id" +
          slotCount +
          ").\nUse fewer rows (max " +
          slotCount +
          ")."
      );
      return null;
    }

    return {
      dataPath: fileEt.text,
      rows: parsed.rows,
      slotCount: slotCount
    };
  }

  function showBatchSummary(lines) {
    var dlg = new Window("dialog", "Sheet loaded");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.margins = 16;
    dlg.preferredSize.width = 520;

    var body = dlg.add("edittext", undefined, lines.join("\n"), {
      multiline: true,
      scrolling: true,
      readonly: true
    });
    body.preferredSize.width = 490;
    body.preferredSize.height = 360;

    dlg.add("statictext", undefined, "Check the PSD. Save if correct.", { multiline: true });
    var btns = dlg.add("group");
    btns.alignment = "right";
    btns.add("button", undefined, "Done", { name: "ok" });
    dlg.show();
  }

  function trim(s) {
    return String(s == null ? "" : s).replace(/^\s+|\s+$/g, "");
  }

  // 1-based data rows after the header. Comma list of singles and a-b ranges.
  // "102-104" → 102,103,104; "102,103,107"; "102-104,107"; reversed "104-102"
  // expands ascending. Typed order is kept; duplicates are dropped.
  function parseRowList(text) {
    var s = trim(String(text == null ? "" : text));
    if (!s) return { error: "empty" };
    var parts = s.split(",");
    var rows = [];
    var seen = {};
    var i;

    function addRow(n) {
      if (seen[n]) return true;
      seen[n] = true;
      rows.push(n);
      return true;
    }

    for (i = 0; i < parts.length; i++) {
      var item = trim(parts[i]);
      if (!item) return { error: "invalid" };
      var single = item.match(/^(\d+)$/);
      if (single) {
        var n = parseInt(single[1], 10);
        if (isNaN(n) || n < 1) return { error: "invalid" };
        addRow(n);
        continue;
      }
      var pair = item.match(/^(\d+)\s*-\s*(\d+)$/);
      if (pair) {
        var a = parseInt(pair[1], 10);
        var b = parseInt(pair[2], 10);
        if (isNaN(a) || isNaN(b) || a < 1 || b < 1) return { error: "invalid" };
        var lo = a <= b ? a : b;
        var hi = a <= b ? b : a;
        var r;
        for (r = lo; r <= hi; r++) addRow(r);
        continue;
      }
      return { error: "invalid" };
    }
    if (!rows.length) return { error: "empty" };
    return { rows: rows };
  }

  function countIdGroups(doc) {
    var found = 0;
    var seen = {};
    if (!doc) return MAX_SLOTS;
    try {
      walkLayers(doc, function (layer) {
        if (!layer || layer.typename !== "LayerSet") return;
        var m = String(layer.name || "").match(/^id\s*(\d+)$/i);
        if (!m) return;
        var n = parseInt(m[1], 10);
        if (isNaN(n) || n < 1 || seen[n]) return;
        seen[n] = true;
        found++;
      });
    } catch (eCount) {
      return MAX_SLOTS;
    }
    return found > 0 ? found : MAX_SLOTS;
  }

  function stripLeadingBom(text) {
    var s = String(text == null ? "" : text);
    if (s.length && s.charCodeAt(0) === 0xfeff) s = s.substring(1);
    // UTF-8 BOM bytes if the file was read as Latin-1 / SYSTEM
    if (
      s.length >= 3 &&
      s.charCodeAt(0) === 0xef &&
      s.charCodeAt(1) === 0xbb &&
      s.charCodeAt(2) === 0xbf
    ) {
      s = s.substring(3);
    }
    if (s.substring(0, 3) === "\u00ef\u00bb\u00bf") s = s.substring(3);
    return s;
  }

  function stripBomFromRows(rows) {
    if (rows && rows.length && rows[0] && rows[0].length) {
      rows[0][0] = stripLeadingBom(rows[0][0]);
    }
    return rows;
  }

  function formatSkip(csvRowNum, column, value, reason) {
    var shown = value == null ? "" : String(value);
    if (shown.length > 120) shown = shown.substring(0, 120) + "...";
    return (
      "row " +
      csvRowNum +
      ", column " +
      (column == null ? "" : String(column)) +
      ", value '" +
      shown +
      "' -- " +
      (reason == null ? "" : String(reason))
    );
  }

  function fileExt(path) {
    var s = String(path || "").toLowerCase();
    var i = s.lastIndexOf(".");
    return i >= 0 ? s.substring(i) : "";
  }

  function looksLikeInvalidUtf8(text) {
    var s = String(text == null ? "" : text);
    if (!s) return false;
    if (s.indexOf("\uFFFD") >= 0) return true;
    // Typical mojibake from invalid or mis-decoded UTF-8 (CP1252 leftover).
    var hints = [
      "Ã±",
      "Ã‘",
      "Ã¡",
      "Ã©",
      "Ã­",
      "Ã³",
      "Ãº",
      "Ã¼",
      "Ã¤",
      "Ã¶",
      "Â¿",
      "Â¡",
      "â€œ",
      "â€˜",
      "â€“",
      "â€”",
      "Ãƒ"
    ];
    var i;
    for (i = 0; i < hints.length; i++) {
      if (s.indexOf(hints[i]) >= 0) return true;
    }
    return false;
  }

  function tryReadFileEncoding(path, encoding) {
    var f = new File(path);
    var text = "";
    try {
      f.encoding = encoding;
      f.open("r");
      text = f.read();
      f.close();
    } catch (eRead) {
      try {
        f.close();
      } catch (eClose) {}
      text = "";
    }
    return text;
  }

  function readFileWithDetectedEncoding(path) {
    var f = new File(path);
    if (!f.exists) throw new Error("File not found:\n" + path);

    // Ken has UTF-8 CSVs (Sheets/Excel) and older Windows-1252 lists, both with ñ.
    var utfNames = ["UTF-8", "UTF8"];
    var sysNames = ["SYSTEM", "Windows-1252", "CP1252"];
    var text = "";
    var encodingUsed = "";
    var i;
    for (i = 0; i < utfNames.length; i++) {
      text = tryReadFileEncoding(path, utfNames[i]);
      if (text && text.length) {
        encodingUsed = "UTF-8";
        break;
      }
    }
    if (text && looksLikeInvalidUtf8(text)) {
      var fallback = "";
      for (i = 0; i < sysNames.length; i++) {
        fallback = tryReadFileEncoding(path, sysNames[i]);
        if (fallback && fallback.length) {
          text = fallback;
          encodingUsed = "SYSTEM (Windows-1252)";
          break;
        }
      }
    }
    if (!text) throw new Error("Could not read file:\n" + path);
    return { text: stripLeadingBom(text), encoding: encodingUsed || "UTF-8" };
  }

  function readFile(path) {
    return readFileWithDetectedEncoding(path).text;
  }

  // Integer → digit string without scientific notation (ExtendScript-safe)
  function integerToDigits(n) {
    if (typeof n !== "number" || isNaN(n)) return "";
    var neg = n < 0;
    n = Math.abs(Math.floor(n + 0.5));
    if (n === 0) return "0";
    var s = "";
    while (n > 0) {
      s = String(n % 10) + s;
      n = Math.floor(n / 10);
    }
    return neg ? "-" + s : s;
  }

  // Read cell as displayed/full text — critical for LRN and local paths
  function excelCellText(cell) {
    // Hyperlink cells: use Address (local path / URL), not display text
    try {
      if (cell.Hyperlinks && cell.Hyperlinks.Count > 0) {
        var hl = cell.Hyperlinks(1);
        var addr = trim(hl.Address || "");
        var sub = "";
        try {
          sub = trim(hl.SubAddress || "");
        } catch (eSub) {}
        if (addr) {
          if (/^file:/i.test(addr)) addr = stripFileUrl(addr);
          return addr;
        }
        if (sub) return sub;
      }
    } catch (eH) {}

    var t = "";
    try {
      t = trim(cell.Text);
    } catch (eT) {
      t = "";
    }
    if (t && !/[eE][+\-]?\d+$/.test(t)) return t;

    var v = null;
    try {
      v = cell.Value2;
    } catch (eV) {
      v = null;
    }

    if (typeof v === "number" && !isNaN(v)) {
      if (Math.abs(v) < 1e15 && Math.abs(v - Math.floor(v + 0.5)) < 1e-6) {
        return integerToDigits(v);
      }
      return String(v);
    }
    if (v != null && v !== undefined) return trim(String(v));
    return t;
  }

  function stripFileUrl(url) {
    var s = trim(url);
    s = s.replace(/^file:\/\/\//i, "");
    s = s.replace(/^file:\/\//i, "");
    try {
      s = decodeURI(s);
    } catch (eD) {}
    return s.replace(/\//g, "\\");
  }

  // Needs Microsoft Excel OR Python openpyxl fallback (excel_to_rows.py)
  function parseExcel(path) {
    try {
      return parseExcelViaCOM(path);
    } catch (eCom) {
      try {
        return parseExcelViaPython(path);
      } catch (ePy) {
        throw new Error(
          "Cannot read Excel file.\n\n" +
            "Excel COM: " + eCom.message + "\n\n" +
            "Python: " + ePy.message + "\n\n" +
            "Fix: install Microsoft Excel, OR keep AccountingSystem .venv (openpyxl),\n" +
            "OR export a CSV with LRN column set to Plain Text."
        );
      }
    }
  }

  function parseExcelViaCOM(path) {
    var excel = null;
    var wb = null;
    try {
      excel = new ActiveXObject("Excel.Application");
    } catch (eApp) {
      throw new Error("Excel not installed / COM blocked. " + eApp.message);
    }

    excel.Visible = false;
    excel.DisplayAlerts = false;

    try {
      var absPath = File(path).fsName;
      wb = excel.Workbooks.Open(absPath, 0, true);
      var sheet = wb.Worksheets(1);
      var used = sheet.UsedRange;
      var rowCount = used.Rows.Count;
      var colCount = used.Columns.Count;
      var startRow = used.Row;
      var startCol = used.Column;

      var rows = [];
      for (var r = 0; r < rowCount; r++) {
        var row = [];
        var any = false;
        for (var c = 0; c < colCount; c++) {
          var cell = sheet.Cells(startRow + r, startCol + c);
          var val = excelCellText(cell);
          if (trim(val)) any = true;
          row.push(val);
        }
        if (any || r === 0) rows.push(row);
      }

      if (rows.length < 2) {
        throw new Error("Excel sheet needs a header row plus at least one data row.");
      }
      return rows;
    } finally {
      try {
        if (wb) wb.Close(false);
      } catch (eClose) {}
      try {
        if (excel) excel.Quit();
      } catch (eQuit) {}
    }
  }

  function scriptFolder() {
    try {
      return File($.fileName).parent;
    } catch (e) {
      return Folder.desktop;
    }
  }

  function findPythonExe() {
    var candidates = [
      "C:/Users/hp/CURSOR/Projects/AccountingSystem/.venv/Scripts/python.exe",
      Folder.myDocuments.parent.fsName + "/CURSOR/Projects/AccountingSystem/.venv/Scripts/python.exe"
    ];
    for (var i = 0; i < candidates.length; i++) {
      var f = new File(candidates[i]);
      if (f.exists) return f.fsName;
    }
    return "py";
  }

  function unescapeTsvCell(s) {
    var out = "";
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      if (ch === "\\" && i + 1 < s.length) {
        var n = s.charAt(i + 1);
        if (n === "n") {
          out += "\n";
          i++;
        } else if (n === "t") {
          out += "\t";
          i++;
        } else if (n === "\\") {
          out += "\\";
          i++;
        } else {
          out += ch;
        }
      } else {
        out += ch;
      }
    }
    return out;
  }

  function parseTSV(text) {
    text = stripLeadingBom(String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n"));
    var lines = text.split("\n");
    var rows = [];
    for (var i = 0; i < lines.length; i++) {
      if (!trim(lines[i])) continue;
      var parts = lines[i].split("\t");
      var row = [];
      for (var c = 0; c < parts.length; c++) row.push(unescapeTsvCell(parts[c]));
      rows.push(row);
    }
    return rows;
  }

  function parseExcelViaPython(path) {
    var converter = new File(scriptFolder().fsName + "/excel_to_rows.py");
    if (!converter.exists) {
      throw new Error("Missing excel_to_rows.py next to batch_fill.jsx");
    }

    var outFile = new File(Folder.temp.fsName + "/ps_id_batch_rows.tsv");
    if (outFile.exists) outFile.remove();

    var py = findPythonExe();
    var cmd =
      '"' +
      py +
      '" "' +
      converter.fsName +
      '" "' +
      File(path).fsName +
      '" "' +
      outFile.fsName +
      '"';

    var code = 0;
    try {
      code = app.system(cmd);
    } catch (eSys) {
      throw new Error("Could not run Python converter.\n" + eSys.message + "\n" + cmd);
    }

    if (!outFile.exists) {
      throw new Error(
        "Python converter failed (exit " +
          code +
          "). Install openpyxl in that Python, or install Excel.\n" +
          cmd
      );
    }

    var rows = parseTSV(readFile(outFile.fsName));
    if (rows.length < 2) {
      throw new Error("Converted Excel has no data rows.");
    }
    return rows;
  }

  function loadTable(path) {
    var ext = fileExt(path);
    if (ext === ".xlsx" || ext === ".xls" || ext === ".xlsm") {
      return { rows: stripBomFromRows(parseExcel(path)), source: "Excel", encoding: "Excel" };
    }
    if (ext === ".tsv") {
      var tsvDetected = readFileWithDetectedEncoding(path);
      return {
        rows: stripBomFromRows(parseTSV(tsvDetected.text)),
        source: "TSV",
        encoding: tsvDetected.encoding
      };
    }
    var detected = readFileWithDetectedEncoding(path);
    var raw = parseCSV(detected.text);
    var repaired = stripBomFromRows(coalesceBrokenCsvRows(raw));
    return {
      rows: repaired,
      source: "CSV",
      encoding: detected.encoding,
      rawRowCount: raw.length,
      repairedRowCount: repaired.length
    };
  }

  // Full-file CSV parse — keeps newlines that sit inside quoted cells
  // (Google Sheets often exports Merits/Affiliation with line breaks).
  function parseCSV(text) {
    text = stripLeadingBom(String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n"));
    var rows = [];
    var row = [];
    var cur = "";
    var inQuotes = false;

    for (var i = 0; i < text.length; i++) {
      var ch = text.charAt(i);
      if (ch === '"') {
        if (inQuotes && i + 1 < text.length && text.charAt(i + 1) === '"') {
          cur += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (ch === "," && !inQuotes) {
        row.push(cur);
        cur = "";
      } else if (ch === "\n" && !inQuotes) {
        row.push(cur);
        cur = "";
        if (row.length > 1 || trim(row[0])) rows.push(row);
        row = [];
      } else {
        cur += ch;
      }
    }

    row.push(cur);
    if (row.length > 1 || trim(row[0])) rows.push(row);
    return rows;
  }

  // Merits/notes often have newlines WITHOUT quotes → fake 1-cell rows.
  // Wide sheets (100+ cols) often OMIT trailing empty cells — those are still
  // valid rows and must NOT be merged into the previous row.
  function coalesceBrokenCsvRows(rows) {
    if (!rows || rows.length < 2) return rows;
    var expected = rows[0].length;
    if (expected < 5) return rows;

    var out = [rows[0].slice()];
    var i;
    for (i = 1; i < rows.length; i++) {
      var row = rows[i].slice();
      var prev = out[out.length - 1];

      // Previous row still missing columns → this line continues a quoted/broken cell
      if (prev.length < expected) {
        prev[prev.length - 1] =
          String(prev[prev.length - 1]) +
          "\n" +
          String(row[0] != null ? row[0] : "");
        for (var j = 1; j < row.length; j++) prev.push(row[j]);
        continue;
      }

      // Only merge true 1-cell orphans (unquoted linebreak mid-cell).
      // Do NOT use a % threshold — that breaks wide CSVs with trimmed trailing commas.
      if (row.length === 1 && expected >= 8) {
        var only = trim(row[0]);
        // New record usually starts with card digits / name — don't swallow those
        var looksNew =
          /^\d{4}[\s-]?\d{4}/.test(only) ||
          /^[A-Za-z][A-Za-z\s.'-]{2,},/.test(only);
        if (!looksNew) {
          prev[prev.length - 1] = String(prev[prev.length - 1]) + "\n" + only;
          continue;
        }
      }

      out.push(row);
    }

    // Pad short rows to header width (trailing empty cols omitted by Sheets/Excel)
    for (i = 1; i < out.length; i++) {
      while (out[i].length < expected) out[i].push("");
      // Keep extras if a row is longer (misaligned export) — don't silently drop paths
    }
    return out;
  }

  function csvCell(cells, headerInfo, column, dataRows, aliases) {
    var idx = resolveColumnIndex(
      headerInfo,
      { column: column, columnAliases: aliases || [] },
      dataRows
    );
    if (idx === undefined) return "";
    return trim(cells && cells[idx] != null ? cells[idx] : "");
  }

  // firstName + lastName, else NAME, else DataB, else DataA
  function rowDisplayName(cells, headerInfo, dataRows) {
    var fn = csvCell(cells, headerInfo, "firstName", dataRows, ["First Name"]);
    var ln = csvCell(cells, headerInfo, "lastName", dataRows, ["Last Name"]);
    var combined = trim(fn + (fn && ln ? " " : "") + ln);
    if (combined) return combined;
    var name = csvCell(cells, headerInfo, "NAME", dataRows, ["Name", "MEMBER NAME"]);
    if (name) return name;
    var dataB = csvCell(cells, headerInfo, "DataB", dataRows);
    if (dataB) return dataB;
    var dataA = csvCell(cells, headerInfo, "DataA", dataRows);
    if (dataA) return dataA;
    return "(no name)";
  }

  // Valid if it has a name (or firstName/lastName), Card Number, or DataA
  function isValidCardRow(cells, headerInfo, dataRows) {
    if (!cells) return false;
    var name = csvCell(cells, headerInfo, "NAME", dataRows, ["Name", "MEMBER NAME"]);
    var fn = csvCell(cells, headerInfo, "firstName", dataRows, ["First Name"]);
    var ln = csvCell(cells, headerInfo, "lastName", dataRows, ["Last Name"]);
    var card = csvCell(cells, headerInfo, "Card Number", dataRows, [
      "CardNumber",
      "CARD NUMBER"
    ]);
    var dataA = csvCell(cells, headerInfo, "DataA", dataRows);
    return !!(name || fn || ln || card || dataA);
  }

  function fileStem(path) {
    var s = trim(path);
    if (!s) return "";
    s = s.replace(/\\/g, "/");
    var slash = s.lastIndexOf("/");
    if (slash >= 0) s = s.substring(slash + 1);
    var dot = s.lastIndexOf(".");
    if (dot > 0) s = s.substring(0, dot);
    return s;
  }

  // Warn (do not fail) when PictureB filename does not match this row's DataA
  function warnIfPictureBMismatch(cells, headerInfo, dataRows, csvRowNum, log) {
    if (!log.warnings) log.warnings = [];
    var dataA = csvCell(cells, headerInfo, "DataA", dataRows);
    var picB = csvCell(cells, headerInfo, "PictureB", dataRows, ["Picture B"]);
    if (!dataA || !picB) return;
    var stem = fileStem(picB);
    if (!stem) return;
    var codeKey = normalizeHeader(dataA);
    var stemKey = normalizeHeader(stem);
    if (!codeKey || stemKey === codeKey) return;
    if (stemKey.indexOf(codeKey) >= 0) return;
    log.warnings.push(
      "CSV row " +
        csvRowNum +
        " PictureB file \"" +
        stem +
        "\" does not match DataA \"" +
        dataA +
        "\""
    );
  }

  function filterStudentRows(headerRow, dataRows, headerInfo) {
    var kept = [];
    var dropped = 0;

    for (var i = 0; i < dataRows.length; i++) {
      if (isValidCardRow(dataRows[i], headerInfo, dataRows)) {
        kept.push(dataRows[i]);
      } else {
        dropped++;
      }
    }
    return { rows: kept, dropped: dropped, minCols: headerRow.length };
  }

  function looksLikeImagePath(s) {
    s = trim(s);
    if (!s) return false;
    if (/^(PRINTED|ENROLLED|TRUE|FALSE|N\/A|NONE)$/i.test(s)) return false;
    if (/\.(png|jpe?g|gif|tif{1,2}|bmp|psd|webp)$/i.test(s)) return true;
    if (/^[A-Za-z]:[\\\/]/.test(s)) return true;
    if (/^\\\\/.test(s)) return true;
    return false;
  }

  // If "qr path" cell is blank/wrong (CSV column shift), find a path that looks like QR
  function recoverImagePath(cells, colIdx, fieldPrefix, pictureColIdx) {
    var primary = colIdx !== undefined && cells[colIdx] != null ? trim(cells[colIdx]) : "";
    if (looksLikeImagePath(primary)) return normalizeLocalPath(primary);

    var i;
    // Prefer nearby columns first (AC±1)
    var near = [];
    if (colIdx !== undefined) {
      near.push(colIdx - 1, colIdx + 1, colIdx + 2, colIdx - 2);
    }
    for (i = 0; i < near.length; i++) {
      var n = near[i];
      if (n < 0 || n >= cells.length) continue;
      if (pictureColIdx !== undefined && n === pictureColIdx) continue;
      var cand = trim(cells[n] != null ? cells[n] : "");
      if (!looksLikeImagePath(cand)) continue;
      if (fieldPrefix === "qr") {
        if (/[\\\/]qr[\\\/]/i.test(cand) || /qr/i.test(cand)) return normalizeLocalPath(cand);
      } else {
        return normalizeLocalPath(cand);
      }
    }

    if (fieldPrefix === "qr") {
      for (i = 0; i < cells.length; i++) {
        if (pictureColIdx !== undefined && i === pictureColIdx) continue;
        var c2 = trim(cells[i] != null ? cells[i] : "");
        if (looksLikeImagePath(c2) && /[\\\/]qr[\\\/]/i.test(c2)) return normalizeLocalPath(c2);
      }
    }

    return normalizeLocalPath(primary);
  }

  function normalizeLocalPath(path) {
    var s = trim(path);
    if (!s) return "";
    s = s.replace(/^"|"$/g, "");
    if (/^file:/i.test(s)) s = stripFileUrl(s);
    return s.replace(/\\/g, "/");
  }

  function decodedPathVariants(path) {
    var out = [];
    function add(p) {
      if (!p) return;
      var s = String(p);
      var i;
      for (i = 0; i < out.length; i++) {
        if (out[i] === s) return;
      }
      out.push(s);
    }
    add(path);
    try {
      add(decodeURI(path));
    } catch (e1) {}
    try {
      add(decodeURIComponent(path));
    } catch (e2) {}
    try {
      if (typeof File !== "undefined" && File.decode) add(File.decode(path));
    } catch (e3) {}
    add(String(path).replace(/%20/g, " "));
    return out;
  }

  // Follow the CSV path exactly (plus slash / quote / missing-.png / URI-decode)
  function resolveImageFile(path) {
    var raw = trim(path);
    if (!raw) return { file: null, tried: [], reason: "empty path" };

    var tried = [];
    var candidates = [];

    function addCandidate(p) {
      if (!p) return;
      for (var i = 0; i < candidates.length; i++) {
        if (candidates[i] === p) return;
      }
      candidates.push(p);
    }

    var bases = decodedPathVariants(raw);
    var bi;
    for (bi = 0; bi < bases.length; bi++) {
      var cleaned = normalizeLocalPath(bases[bi]);
      var back = cleaned.replace(/\//g, "\\");
      addCandidate(cleaned);
      addCandidate(back);

      // Same path if CSV dropped the extension
      if (!/\.(png|jpe?g|gif|tif{1,2}|bmp|webp)$/i.test(back)) {
        addCandidate(back + ".png");
        addCandidate(cleaned + ".png");
      }
    }

    var i;
    for (i = 0; i < candidates.length; i++) {
      var p = candidates[i];
      tried.push(p);
      try {
        var f = new File(p);
        if (f.exists) return { file: prepareImageFile(f), tried: tried, reason: "" };
      } catch (eOpen) {}
    }

    var display = normalizeLocalPath(raw).replace(/\//g, "\\");
    return {
      file: null,
      tried: tried,
      reason: "file not found at CSV path:\n" + display
    };
  }

  // ExtendScript/CS6 often fails to open paths with "=" or very long names — copy to temp
  function prepareImageFile(srcFile) {
    var name = String(srcFile.name);
    var needsCopy =
      name.indexOf("=") >= 0 ||
      name.indexOf("?") >= 0 ||
      name.length > 80 ||
      name.indexOf("..") >= 0;

    if (!needsCopy) return srcFile;

    var ext = ".png";
    var m = name.match(/(\.[A-Za-z0-9]+)$/);
    if (m) ext = m[1].toLowerCase();

    var token = "";
    var tm = name.match(/token=([A-Za-z0-9_-]+)/i);
    if (tm) token = tm[1];
    if (!token) token = String(new Date().getTime());

    var dest = new File(Folder.temp.fsName + "/ps_qr_" + token + ext);
    try {
      if (dest.exists) dest.remove();
    } catch (eRm) {}

    try {
      if (srcFile.copy(dest)) return dest;
    } catch (eCopy) {}

    // Fallback: cmd copy
    try {
      app.system(
        'cmd /c copy /Y "' + srcFile.fsName + '" "' + dest.fsName + '"'
      );
      if (dest.exists) return dest;
    } catch (eCmd) {}

    return srcFile;
  }

  // Match real QR file in folder by token=XXXX (CSV path is often truncated)
  function findQrFileInFolder(fullPath) {
    var back = String(fullPath).replace(/\//g, "\\");
    back = back.replace(/\.(png|jpe?g|gif|tif{1,2}|bmp|webp)$/i, "");
    var slash = Math.max(back.lastIndexOf("\\"), back.lastIndexOf("/"));
    if (slash < 0) return null;

    var dir = back.substring(0, slash);
    var base = back.substring(slash + 1);
    var folder = new Folder(dir);
    if (!folder.exists) {
      folder = new Folder(dir.replace(/\\/g, "/"));
    }
    if (!folder.exists) return null;

    var token = "";
    var m = base.match(/token=([A-Za-z0-9_-]+)/i);
    if (m) token = m[1];
    if (!token) {
      var m2 = base.match(new RegExp("=([A-Za-z0-9_-]{6,})$"));
      if (m2) token = m2[1];
    }
    if (!token) return null;

    var tokenLower = token.toLowerCase();
    var files = null;

    try {
      files = folder.getFiles("*" + token + "*");
    } catch (eWild) {
      files = null;
    }

    if (!files || !files.length) {
      try {
        files = folder.getFiles("*.png");
      } catch (ePng) {
        files = [];
      }
      try {
        var jpgs = folder.getFiles("*.jpg");
        if (jpgs && jpgs.length) files = files.concat(jpgs);
      } catch (eJpg) {}
    }

    if (!files || !files.length) return null;

    var i;
    var best = null;
    for (i = 0; i < files.length; i++) {
      if (!(files[i] instanceof File)) continue;
      var name = String(files[i].name).toLowerCase();
      if (name.indexOf(tokenLower) >= 0) {
        best = files[i];
        break;
      }
    }
    return best;
  }

  // Case-insensitive; spaces, underscores, and hyphens ignored
  function normalizeHeader(s) {
    return trim(String(s == null ? "" : s))
      .toLowerCase()
      .replace(/[\s_\-]+/g, "");
  }

  // Index every header; keep ALL duplicate positions (wide sheets often repeat names)
  function buildHeaderInfo(headerCells) {
    var byNorm = {};
    var originals = [];
    var dupes = [];
    for (var i = 0; i < headerCells.length; i++) {
      var raw = trim(headerCells[i]);
      originals[i] = raw;
      if (!raw) continue;
      var n = normalizeHeader(raw);
      if (!byNorm[n]) byNorm[n] = [];
      if (byNorm[n].length) {
        dupes.push(raw + " @ " + excelColLetter(i) + " (also " + excelColLetter(byNorm[n][0]) + ")");
      }
      byNorm[n].push(i);
    }
    return {
      byNorm: byNorm,
      originals: originals,
      width: headerCells.length,
      __dupes: dupes
    };
  }

  function uniqueInts(arr) {
    var out = [];
    for (var i = 0; i < arr.length; i++) {
      var v = arr[i];
      var seen = false;
      for (var j = 0; j < out.length; j++) {
        if (out[j] === v) {
          seen = true;
          break;
        }
      }
      if (!seen) out.push(v);
    }
    return out;
  }

  // If duplicate headers, prefer the column that actually has values / image paths
  function pickBestColIndex(indices, dataRows, preferPaths) {
    if (!indices || !indices.length) return undefined;
    if (indices.length === 1) return indices[0];

    var best = indices[0];
    var bestScore = -1;
    var limit = Math.min(dataRows ? dataRows.length : 0, 40);

    for (var i = 0; i < indices.length; i++) {
      var idx = indices[i];
      var score = 0;
      for (var r = 0; r < limit; r++) {
        var v =
          dataRows[r] && dataRows[r][idx] != null ? trim(dataRows[r][idx]) : "";
        if (!v) continue;
        if (preferPaths && looksLikeImagePath(v)) score += 10;
        else score += 1;
      }
      // Prefer later columns on tie (often the real path col on wide sheets)
      if (score > bestScore || (score === bestScore && idx > best)) {
        bestScore = score;
        best = idx;
      }
    }
    return best;
  }

  function resolveColumnIndex(headerInfo, field, dataRows) {
    if (!field || !headerInfo) return undefined;

    var candidates = [];

    function addNorm(name) {
      if (!name) return;
      var list = headerInfo.byNorm[normalizeHeader(name)];
      if (!list) return;
      for (var i = 0; i < list.length; i++) candidates.push(list[i]);
    }

    addNorm(field.column);
    var aliases = field.columnAliases || [];
    for (var a = 0; a < aliases.length; a++) addNorm(aliases[a]);

    // smartA / dataA → also match PictureA-style aliases already listed on field
    var m = String(field.prefix || "").match(/^(smart|data)([A-Ta-t])$/i);
    if (m) {
      var kind = m[1].toLowerCase();
      var letter = m[2].toUpperCase();
      if (kind === "smart") {
        addNorm("Picture" + letter);
        addNorm("Picture " + letter);
      } else {
        addNorm("Data" + letter);
        addNorm("Data " + letter);
      }
      addNorm(kind + letter);
      addNorm(kind + " " + letter);
    }

    candidates = uniqueInts(candidates);
    return pickBestColIndex(
      candidates,
      dataRows || [],
      field.type === "image"
    );
  }

  // Compat: old col["NAME"] style — exact then normalized
  function colGet(headerInfo, exactName, dataRows) {
    return resolveColumnIndex(
      headerInfo,
      { column: exactName, type: "text" },
      dataRows
    );
  }

  function excelColLetter(idx) {
    var n = idx + 1;
    var s = "";
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function repeatChar(ch, count) {
    var out = "";
    for (var i = 0; i < count; i++) out += ch;
    return out;
  }

  // Fixes CSV values like 1.27272E+11 → 127272260007
  function formatLRN(raw) {
    var s = trim(raw);
    if (!s) return "";

    // Already plain digits
    if (/^\d+$/.test(s)) return s;

    // Scientific notation from Excel/Sheets export
    var sci = s.match(/^([+-]?)(\d+(?:\.\d+)?)[eE]([+-]?\d+)$/);
    if (sci) {
      var sign = sci[1];
      var mantissa = sci[2];
      var exp = parseInt(sci[3], 10);
      var parts = mantissa.split(".");
      var intPart = parts[0];
      var decPart = parts.length > 1 ? parts[1] : "";
      var digits = intPart + decPart;
      var newLen = intPart.length + exp;

      if (newLen <= 0) return "0";

      if (newLen > digits.length) {
        digits = digits + repeatChar("0", newLen - digits.length);
      } else if (newLen < digits.length) {
        digits = digits.substring(0, newLen);
      }

      digits = digits.replace(/^0+/, "");
      if (!digits) digits = "0";
      return sign === "-" ? "-" + digits : digits;
    }

    // Plain decimal number: 127272260007.0
    if (/^\d+\.\d+$/.test(s)) {
      var numParts = s.split(".");
      if (/^0+$/.test(numParts[1])) return numParts[0];
    }

    // Strip commas/spaces, keep digits only if nothing else
    var digitsOnly = s.replace(/[, \t]/g, "");
    if (/^\d+$/.test(digitsOnly)) return digitsOnly;

    return s;
  }

  // Card Number: 8155526530903026 → 8155 5265 3090 3026 (same as sheet paste)
  function formatCardNumber(raw) {
    var s = trim(raw);
    if (!s) return "";
    var digits = s.replace(/\D/g, "");
    if (!digits) return s;
    if (digits.length === 16) {
      return (
        digits.substring(0, 4) +
        " " +
        digits.substring(4, 8) +
        " " +
        digits.substring(8, 12) +
        " " +
        digits.substring(12, 16)
      );
    }
    // Already spaced correctly in sheet — keep spaces as typed if 19 chars with spaces
    if (/^\d{4}\s\d{4}\s\d{4}\s\d{4}$/.test(s)) return s;
    return s;
  }

  // Keep leading zeros, no spaces (QR Card Number)
  function formatDigits(raw) {
    var s = trim(raw);
    if (!s) return "";
    var digits = s.replace(/\D/g, "");
    return digits || s;
  }

  function formatContact(raw) {
    var s = trim(raw);
    if (!s) return "";
    var upper = s.toUpperCase();
    if (upper.indexOf("W/O") >= 0 || upper.indexOf("N/A") >= 0 || upper.indexOf("NONE") >= 0) {
      return s;
    }

    var digits = s.replace(/\D/g, "");
    if (!digits) return s;

    if (digits.charAt(0) !== "0") {
      if (digits.length === 10) digits = "0" + digits;
      else if (digits.length === 11 && digits.charAt(0) === "6") {
        digits = "0" + digits.substring(1);
      }
    }

    if (digits.length === 11 && digits.charAt(0) === "0") {
      return (
        digits.substring(0, 4) +
        " " +
        digits.substring(4, 7) +
        " " +
        digits.substring(7, 11)
      );
    }

    return digits;
  }

  function cellValue(cells, colIndex, fieldType) {
    if (colIndex === undefined) return null;
    var v = trim(cells[colIndex] != null ? cells[colIndex] : "");
    if (fieldType === "lrn") return formatLRN(v);
    if (fieldType === "cardnumber") return formatCardNumber(v);
    if (fieldType === "digits") return formatDigits(v);
    if (fieldType === "contact") return formatContact(v);
    if (fieldType === "image") return normalizeLocalPath(v);
    return v;
  }

  function nameKey(s) {
    return trim(s).toLowerCase().replace(/\s+/g, "");
  }

  function walkLayers(container, fn) {
    for (var i = 0; i < container.layers.length; i++) {
      var layer = container.layers[i];
      fn(layer);
      if (layer.typename === "LayerSet") walkLayers(layer, fn);
    }
  }

  function findGroupByName(container, name) {
    var want = nameKey(name);
    var found = null;
    walkLayers(container, function (layer) {
      if (found) return;
      if (layer.typename === "LayerSet" && nameKey(layer.name) === want) found = layer;
    });
    return found;
  }

  function findLayerForSlot(doc, slot, fieldName, aliases) {
    var wantGroup = nameKey("id" + slot);
    // data1 / smart1 already include a number — do NOT append slot (avoids smart11)
    var exactOnly = /\d+$/.test(String(fieldName));
    var unnumbered = nameKey(fieldName);
    var numbered = exactOnly ? unnumbered : nameKey(fieldName + slot);
    var extra = {};
    if (aliases) {
      for (var a = 0; a < aliases.length; a++) extra[nameKey(aliases[a])] = true;
    }

    var group = findGroupByName(doc, "id" + slot);
    var inGroupNumbered = null;
    var inGroupUnnumbered = null;
    var inGroupAlias = null;
    var inGroupPrefix = null;
    var anywhereNumbered = null;
    var anywhereExact = null;

    function underWantedGroup(layer) {
      var p = layer.parent;
      while (p && p.typename === "LayerSet") {
        if (nameKey(p.name) === wantGroup) return true;
        try {
          p = p.parent;
        } catch (e) {
          break;
        }
      }
      return false;
    }

    walkLayers(doc, function (layer) {
      if (layer.typename === "LayerSet") return;
      var nk = nameKey(layer.name);
      var inWanted = underWantedGroup(layer);

      if (nk.indexOf("__pin") === 0) return;

      if (nk === numbered) {
        if (inWanted && !inGroupNumbered) inGroupNumbered = layer;
        if (!anywhereNumbered) anywhereNumbered = layer;
      }
      if (nk === unnumbered) {
        if (inWanted && !inGroupUnnumbered) inGroupUnnumbered = layer;
        if (!anywhereExact) anywhereExact = layer;
      }
      if (inWanted && extra[nk] && !inGroupAlias) inGroupAlias = layer;
      // Prefix match only for base names like "smart" / "qr" (not smart1 → smart10)
      if (
        !exactOnly &&
        inWanted &&
        !inGroupPrefix &&
        nk.indexOf(unnumbered) === 0 &&
        nk !== unnumbered &&
        nk !== numbered
      ) {
        inGroupPrefix = layer;
      }
    });

    function hit(layer, foundLabel) {
      var grp = group;
      if (!grp && layer && layer.parent && layer.parent.typename === "LayerSet") {
        // Walk up to id#
        var p = layer.parent;
        while (p && p.typename === "LayerSet") {
          if (nameKey(p.name) === wantGroup) {
            grp = p;
            break;
          }
          try {
            p = p.parent;
          } catch (e2) {
            break;
          }
        }
        if (!grp) grp = layer.parent;
      }
      return { layer: layer, label: foundLabel, group: grp };
    }

    if (inGroupNumbered) return hit(inGroupNumbered, "id" + slot + "/" + fieldName + (exactOnly ? "" : slot));
    if (inGroupUnnumbered) return hit(inGroupUnnumbered, "id" + slot + "/" + fieldName);
    if (inGroupAlias) return hit(inGroupAlias, "id" + slot + "/" + inGroupAlias.name);
    if (inGroupPrefix) return hit(inGroupPrefix, "id" + slot + "/" + inGroupPrefix.name);
    if (anywhereNumbered) return hit(anywhereNumbered, fieldName + (exactOnly ? "" : String(slot)) + " (flat)");
    if (exactOnly && anywhereExact) return hit(anywhereExact, fieldName + " (flat)");

    return {
      layer: null,
      label: "id" + slot + "/" + fieldName + (exactOnly ? "" : " or " + fieldName + slot),
      group: group
    };
  }

  function charCount(text) {
    return String(text == null ? "" : text).length;
  }

  function lerpByChars(n, fullAt, fullVal, anchorAt, anchorVal, minVal, maxVal) {
    if (n <= fullAt) return fullVal;
    var span = anchorAt - fullAt;
    var slope = span > 0 ? (anchorVal - fullVal) / span : 0;
    var val =
      n <= anchorAt
        ? fullVal + slope * (n - fullAt)
        : anchorVal + slope * (n - anchorAt);
    if (minVal != null && val < minVal) val = minVal;
    if (maxVal != null && val > maxVal) val = maxVal;
    return val;
  }

  // Tracking first (negative = tighter). Scale: optional mid point then anchor.
  function fitMetricsForChars(text, rules) {
    var n = charCount(text);
    var track = lerpByChars(
      n,
      rules.fullAt,
      rules.fullTrack,
      rules.anchorAt,
      rules.anchorTrack,
      rules.minTrack,
      0
    );

    var scale;
    if (rules.midAt != null && rules.midScale != null) {
      if (n <= rules.scaleFullAt) {
        scale = rules.scaleFull;
      } else if (n <= rules.midAt) {
        scale = lerpByChars(
          n,
          rules.scaleFullAt,
          rules.scaleFull,
          rules.midAt,
          rules.midScale,
          rules.minScale,
          100
        );
      } else {
        scale = lerpByChars(
          n,
          rules.midAt,
          rules.midScale,
          rules.scaleAnchorAt,
          rules.scaleAnchor,
          rules.minScale,
          100
        );
      }
    } else {
      scale = lerpByChars(
        n,
        rules.scaleFullAt,
        rules.scaleFull,
        rules.scaleAnchorAt,
        rules.scaleAnchor,
        rules.minScale,
        100
      );
    }

    return {
      track: Math.round(track),
      scale: Math.round(scale * 10) / 10
    };
  }

  function applyHorizontalScale(layer, scale) {
    try {
      layer.textItem.horizontalScale = scale;
    } catch (e) {
      layer.resize(scale, 100, AnchorPosition.MIDDLECENTER);
    }
  }

  function applyTracking(layer, tracking) {
    try {
      layer.textItem.tracking = tracking;
    } catch (e) {}
  }

  function setTextLayerContents(layer, value) {
    if (layer.kind !== LayerKind.TEXT) {
      throw new Error("Layer '" + layer.name + "' is not a text layer.");
    }
    layer.textItem.contents = String(value == null ? "" : value);
  }

  function setPointTextWithTrackingFit(layer, value, rules) {
    if (layer.kind !== LayerKind.TEXT) {
      throw new Error("Layer '" + layer.name + "' is not a text layer.");
    }
    var text = String(value == null ? "" : value);
    var m = fitMetricsForChars(text, rules);
    layer.textItem.contents = text;
    applyTracking(layer, m.track);
    applyHorizontalScale(layer, m.scale);
    return m;
  }

  function boundsPx(layer) {
    var b = layer.bounds;
    return {
      left: b[0].as("px"),
      top: b[1].as("px"),
      right: b[2].as("px"),
      bottom: b[3].as("px"),
      width: b[2].as("px") - b[0].as("px"),
      height: b[3].as("px") - b[1].as("px"),
      cx: (b[0].as("px") + b[2].as("px")) / 2,
      cy: (b[1].as("px") + b[3].as("px")) / 2
    };
  }

  function wasClipped(layer) {
    try {
      return layer.grouped === true;
    } catch (e) {
      return false;
    }
  }

  function makeClippingMask() {
    try {
      var desc = new ActionDescriptor();
      var ref = new ActionReference();
      ref.putEnumerated(charIDToTypeID("Lyr "), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
      desc.putReference(charIDToTypeID("null"), ref);
      executeAction(charIDToTypeID("GrpL"), desc, DialogModes.NO);
    } catch (e) {}
  }

  // Select by unique layer id — never by name (many groups share "smart")
  function selectLayer(layer) {
    try {
      var ref = new ActionReference();
      ref.putIdentifier(charIDToTypeID("Lyr "), layer.id);
      var desc = new ActionDescriptor();
      desc.putReference(charIDToTypeID("null"), ref);
      desc.putBoolean(charIDToTypeID("MkVs"), false);
      executeAction(charIDToTypeID("slct"), desc, DialogModes.NO);
      return;
    } catch (eId) {}
    try {
      app.activeDocument.activeLayer = layer;
    } catch (e1) {}
  }

  function closeDocQuietly(opened) {
    if (!opened) return;
    try {
      opened.close(SaveOptions.DONOTSAVECHANGES);
    } catch (eCloseDoc) {}
  }

  // Duplicate into THIS slot's group, directly before the placeholder.
  // File → Place in CS6 often lands one group above (id3 image → id4).
  function duplicateImageBefore(doc, placeholder, imgFile) {
    var src = app.open(imgFile);
    var dup = null;
    try {
      try {
        if (src.layers.length > 1) src.mergeVisibleLayers();
      } catch (eMerge) {
        try { src.flatten(); } catch (eFlat) {}
      }
      dup = src.activeLayer.duplicate(placeholder, ElementPlacement.PLACEBEFORE);
    } finally {
      closeDocQuietly(src);
      app.activeDocument = doc;
    }
    return dup;
  }

  // Scale image to fit INSIDE the PSD placeholder box (same position, no crop)
  function scaleLayerToFit(layer, target) {
    var cur = boundsPx(layer);
    if (cur.width < 1 || cur.height < 1) return;
    var scaleX = (target.width / cur.width) * 100;
    var scaleY = (target.height / cur.height) * 100;
    var scale = Math.min(scaleX, scaleY);
    layer.resize(scale, scale, AnchorPosition.MIDDLECENTER);
    cur = boundsPx(layer);
    layer.translate(target.cx - cur.cx, target.cy - cur.cy);
  }

  function groupNameOf(layer) {
    try {
      if (layer.parent && layer.parent.typename === "LayerSet") return layer.parent.name;
    } catch (e) {}
    return "(root)";
  }

  // Uses smart/qr layer in PSD as the frame — location + size from template
  function fitImageIntoLayer(doc, placeholder, imagePath, expectedGroup, finalName) {
    var resolved = resolveImageFile(imagePath);
    if (!resolved.file) throw new Error(resolved.reason || "Image not found");

    var imgFile = resolved.file;
    var target = boundsPx(placeholder);
    if (target.width < 1 || target.height < 1) {
      throw new Error("Layer '" + placeholder.name + "' has empty bounds.");
    }

    var restoreName = finalName || placeholder.name;
    if (String(restoreName).indexOf("__pin_") === 0) restoreName = finalName || "smart";

    var keepClip = wasClipped(placeholder);
    var pinName = "__pin_" + placeholder.id;
    placeholder.name = pinName;

    var placed;
    try {
      try {
        placed = duplicateImageBefore(doc, placeholder, imgFile);
      } catch (eDup) {
        selectLayer(placeholder);
        var src = app.open(imgFile);
        try {
          src.selection.selectAll();
          src.selection.copy();
        } finally {
          closeDocQuietly(src);
        }
        app.activeDocument = doc;
        selectLayer(placeholder);
        doc.paste();
        placed = doc.activeLayer;
        try {
          placed.move(placeholder, ElementPlacement.PLACEBEFORE);
        } catch (eMove) {}
      }

      if (!placed) throw new Error("Could not import image.");

      try {
        placed.move(placeholder, ElementPlacement.PLACEBEFORE);
      } catch (eFix) {}

      scaleLayerToFit(placed, target);

      selectLayer(placeholder);
      try {
        placeholder.remove();
      } catch (eDel) {
        placeholder.name = restoreName;
      }

      placed.name = restoreName;
      selectLayer(placed);

      if (keepClip) {
        // Don't clip onto a qr layer (would hide/break the QR)
        var belowOk = true;
        try {
          var parent = placed.parent;
          var layers = parent.layers;
          for (var li = 0; li < layers.length - 1; li++) {
            if (layers[li] === placed) {
              var below = layers[li + 1];
              if (below && nameKey(below.name).indexOf("qr") === 0) belowOk = false;
              break;
            }
          }
        } catch (eBelow) {}
        if (belowOk) makeClippingMask();
      }
    } catch (err) {
      try {
        if (placeholder && placeholder.name === pinName) placeholder.name = restoreName;
      } catch (eRest) {}
      throw err;
    }
  }

  function fillSlot(doc, slot, csvRowNum, cells, headerInfo, dataRows, log) {
    var student = rowDisplayName(cells, headerInfo, dataRows);
    var picCol = resolveColumnIndex(
      headerInfo,
      { column: "PictureA", type: "image", columnAliases: ["pictureA"] },
      dataRows
    );

    warnIfPictureBMismatch(cells, headerInfo, dataRows, csvRowNum, log);

    for (var f = 0; f < FIELDS.length; f++) {
      var field = FIELDS[f];
      var value = "";
      try {
        if (field.enabled === false) continue;

        var colIdx = resolveColumnIndex(headerInfo, field, dataRows);
        if (colIdx === undefined) {
          continue;
        }

        value =
          field.type === "image"
            ? recoverImagePath(cells, colIdx, field.prefix, picCol)
            : cellValue(cells, colIdx, field.type);

        if (value == null || value === "") {
          continue;
        }

        var hit = findLayerForSlot(doc, slot, field.prefix, field.layerAliases);
        if (!hit.layer) {
          if (field.optional === true) continue;
          log.skipped.push(
            formatSkip(csvRowNum, field.column, value, "layer not found (" + hit.label + ")")
          );
          continue;
        }

        if (field.type === "image") {
          if (!looksLikeImagePath(value)) {
            log.skipped.push(
              formatSkip(csvRowNum, field.column, value, "not a file path")
            );
            continue;
          }
          var resolvedEarly = resolveImageFile(value);
          if (!resolvedEarly.file) {
            log.skipped.push(
              formatSkip(csvRowNum, field.column, value, "file not found")
            );
            continue;
          }
          // Restore as smartA1 / smartB1 (prefix + slot)
          fitImageIntoLayer(
            doc,
            hit.layer,
            value,
            hit.group,
            field.prefix + slot
          );
          log.applied.push(
            "CSV row " +
              csvRowNum +
              " " +
              hit.label +
              " OK ← " +
              String(value).replace(/^.*[\\\/]/, "")
          );
          continue;
        }

        if (field.type === "text-fit-track") {
          var m = setPointTextWithTrackingFit(hit.layer, value, field.fitRules);
          log.applied.push(
            hit.label +
              " (" +
              charCount(value) +
              "ch track " +
              m.track +
              " / " +
              m.scale +
              "%)"
          );
        } else {
          setTextLayerContents(hit.layer, value);
          log.applied.push(hit.label);
        }
      } catch (eField) {
        if (field.optional === true && field.type !== "image") continue;
        log.skipped.push(
          formatSkip(
            csvRowNum,
            field.column,
            value,
            eField && eField.message ? eField.message : "field error"
          )
        );
      }
    }

    log.lines.push("CSV row " + csvRowNum + " → id" + slot + ": " + student);
  }

  var settings = showDialog();
  if (!settings) return;

  var originalRuler = app.preferences.rulerUnits;
  var originalDialogs = app.displayDialogs;
  app.preferences.rulerUnits = Units.PIXELS;
  app.displayDialogs = DialogModes.NO;

  var doc = app.activeDocument;
  var batchLog = { applied: [], skipped: [], warnings: [], lines: [] };

  try {
    var loaded = loadTable(settings.dataPath);
    var rows = loaded.rows;
    if (rows.length < 2) {
      throw new Error("Data file needs a header row plus at least one data row.");
    }

    var header = rows[0];
    var dataOnly = rows.slice(1);
    var headerInfo = buildHeaderInfo(header);
    var filtered = filterStudentRows(header, dataOnly, headerInfo);
    // Range uses 1-based data rows after the header (Excel/Sheets row 2 = row 1).
    // Blank/junk rows stay in place so they can be skipped and listed.
    rows = [header].concat(dataOnly);

    var mapLines = [
      "Source: " + loaded.source,
      "CSV encoding: " + (loaded.encoding || "n/a"),
      "Columns in file: " + header.length + " (wide sheets OK — match by header name)"
    ];
    if (loaded.rawRowCount && loaded.rawRowCount !== loaded.repairedRowCount) {
      mapLines.push(
        "CSV repair: " +
          loaded.rawRowCount +
          " raw lines → " +
          loaded.repairedRowCount +
          " rows (merged line-breaks in cells)"
      );
    }
    if (filtered.dropped) {
      mapLines.push(
        "Dropped " + filtered.dropped + " incomplete/junk row(s); using real cards only"
      );
    }
    var unusedHeaders = [];
    for (var fi = 0; fi < FIELDS.length; fi++) {
      var fld = FIELDS[fi];
      if (fld.enabled === false) continue;
      var idx = resolveColumnIndex(headerInfo, fld, dataOnly);
      if (idx === undefined) {
        if (fld.optional !== true) unusedHeaders.push(fld.prefix);
        continue;
      }
      var hdrName = headerInfo.originals[idx] || fld.column;
      var sample = "";
      if (fld.type === "image" && dataOnly.length) {
        var sv = dataOnly[0][idx] != null ? trim(dataOnly[0][idx]) : "";
        sample = sv
          ? " sample: " + sv.replace(/^.*[\\\/]/, "").substring(0, 40)
          : " sample: (empty row1)";
      }
      mapLines.push(
        fld.prefix +
          " ← \"" +
          hdrName +
          "\" (" +
          excelColLetter(idx) +
          ")" +
          sample
      );
    }
    if (unusedHeaders.length) {
      mapLines.push("Not in this CSV (OK): " + unusedHeaders.join(", "));
    }
    if (headerInfo.__dupes && headerInfo.__dupes.length) {
      mapLines.push("Duplicate headers (picked best by data): " + headerInfo.__dupes.join("; "));
    }
    batchLog.lines.push(mapLines.join("\n"));
    batchLog.lines.push("");

    var dataRowCount = rows.length - 1;
    if (dataRowCount < 1) {
      throw new Error(
        "No data rows found after CSV repair.\n" +
          "Re-export CSV from Sheets (Download → CSV)."
      );
    }

    var selectedRows = settings.rows;
    var preview = [];
    var pi;
    for (pi = 0; pi < selectedRows.length; pi++) {
      var p = selectedRows[pi];
      if (p > dataRowCount || !rows[p]) {
        preview.push("  " + p + " = (past end of CSV — skipped)");
      } else if (!isValidCardRow(rows[p], headerInfo, dataOnly)) {
        preview.push("  " + p + " = (blank — skipped)");
      } else {
        preview.push("  " + p + " = " + rowDisplayName(rows[p], headerInfo, dataOnly));
      }
    }

    // Image path pre-check (PictureA/PictureB/…) — warn only, never abort
    var pathPreview = [];
    var missingFiles = 0;
    var fi2;
    for (fi2 = 0; fi2 < FIELDS.length; fi2++) {
      var imgField = FIELDS[fi2];
      if (imgField.enabled === false || imgField.type !== "image") continue;
      var pidx = resolveColumnIndex(headerInfo, imgField, dataOnly);
      var colLabel = imgField.column || imgField.prefix;
      if (pidx === undefined) {
        if (imgField.prefix === "smartA") {
          pathPreview.push(
            "  PictureA / smart1path: HEADER NOT FOUND\n" +
              "  Tip: rename column to PictureA, or keep smart1path (also accepted)."
          );
        }
        continue;
      }
      var pr;
      var pri;
      for (pri = 0; pri < selectedRows.length; pri++) {
        pr = selectedRows[pri];
        if (pr > dataRowCount || !rows[pr]) {
          pathPreview.push("  r" + pr + " " + colLabel + " = (past end of CSV — skipped)");
          continue;
        }
        var cv = rows[pr][pidx] != null ? trim(rows[pr][pidx]) : "";
        var exists = false;
        if (cv && looksLikeImagePath(cv)) {
          var rf = resolveImageFile(cv);
          exists = !!(rf && rf.file);
          if (!exists) missingFiles++;
        }
        pathPreview.push(
          "  r" +
            pr +
            " " +
            colLabel +
            " @" +
            excelColLetter(pidx) +
            " (" +
            (headerInfo.originals[pidx] || "") +
            ") = " +
            (cv
              ? cv.replace(/^.*[\\\/]/, "").substring(0, 40) +
                (exists ? " [OK]" : " [FILE MISSING]")
              : "(empty — skipped)")
        );
      }
    }
    if (!pathPreview.length) {
      pathPreview.push("  smartA / PictureA: not configured");
    }

    batchLog.lines.push(
      "Data rows: " +
        dataRowCount +
        " · valid cards: " +
        filtered.rows.length +
        " (" +
        loaded.source +
        ") · encoding: " +
        (loaded.encoding || "n/a") +
        " · filling rows " +
        selectedRows.join(", ") +
        " into id1–id" +
        selectedRows.length
    );
    batchLog.lines.push("Row check:\n" + preview.join("\n"));
    batchLog.lines.push("Image path check:\n" + pathPreview.join("\n"));
    batchLog.lines.push("");
    if (missingFiles > 0) {
      batchLog.warnings.push(
        "Pre-check: " +
          missingFiles +
          " image file(s) not found at the path in the CSV. " +
          "Those fields will be skipped; other cards and fields still fill.\n" +
          pathPreview.join("\n")
      );
    }

    var fillDone = false;
    function runFillSlots() {
      if (fillDone) return;
      fillDone = true;
      var si;
      for (si = 0; si < selectedRows.length; si++) {
        var r = selectedRows[si];
        var cells = rows[r];
        var slot = si + 1;
        if (!cells || r > dataRowCount) {
          batchLog.skipped.push(formatSkip(r, "(row)", "", "past end of CSV"));
          continue;
        }
        if (!isValidCardRow(cells, headerInfo, dataOnly)) {
          batchLog.skipped.push(
            formatSkip(
              r,
              "(row)",
              rowDisplayName(cells, headerInfo, dataOnly),
              "blank / not a card row"
            )
          );
          continue;
        }
        try {
          fillSlot(doc, slot, r, cells, headerInfo, dataOnly, batchLog);
        } catch (eRow) {
          batchLog.skipped.push(
            formatSkip(
              r,
              "(row)",
              rowDisplayName(cells, headerInfo, dataOnly),
              eRow && eRow.message ? eRow.message : "row error"
            )
          );
        }
      }
    }

    // One History step; skip CS6 screen refresh.
    try {
      $.global.__psNftRunFill = runFillSlots;
      doc.suspendHistory("NFT Back Fill Data20", "$.global.__psNftRunFill()");
    } catch (eHist) {
    } finally {
      try {
        $.global.__psNftRunFill = null;
      } catch (eG) {}
    }
    if (!fillDone) runFillSlots();

    app.displayDialogs = DialogModes.ALL;
    var summary = [];
    summary.push("=== RESULT ===");
    summary.push("Applied: " + batchLog.applied.length);
    summary.push("Skipped: " + batchLog.skipped.length);
    if (batchLog.warnings && batchLog.warnings.length) {
      summary.push("");
      summary.push("--- WARNINGS ---");
      summary = summary.concat(batchLog.warnings);
    }
    if (batchLog.skipped.length) {
      summary.push("");
      summary.push("--- SKIPPED / MISSING (row, column, value) ---");
      summary = summary.concat(batchLog.skipped);
    }
    var imgApplied = [];
    for (var ai = 0; ai < batchLog.applied.length; ai++) {
      var al = batchLog.applied[ai];
      if (
        al.indexOf("/data") >= 0 ||
        al.indexOf("/smart") >= 0 ||
        al.indexOf(" data") >= 0 ||
        al.indexOf(" smart") >= 0 ||
        al.indexOf("OK ←") >= 0
      ) {
        imgApplied.push(al);
      }
    }
    if (imgApplied.length) {
      summary.push("");
      summary.push("--- IMAGES OK ---");
      summary = summary.concat(imgApplied);
    }
    summary.push("");
    summary = summary.concat(batchLog.lines);
    showBatchSummary(summary);
  } catch (err) {
    app.displayDialogs = DialogModes.ALL;
    alert("Stopped:\n" + err.message);
  } finally {
    app.preferences.rulerUnits = originalRuler;
    app.displayDialogs = originalDialogs;
  }
})();
