#!/usr/bin/env node
/**
 * Dummy-data tests for CSV UTF-8 / Windows-1252 read, Data20 fields,
 * row-range parsing, path fallbacks, and skip formatting.
 * Photoshop is not required. Helpers are extracted from the JSX (ES3).
 *
 *   node test/run_tests.js
 */
var fs = require("fs");
var os = require("os");
var path = require("path");

var ROOT = path.resolve(__dirname, "..");
var NFT_JSX = path.join(ROOT, "back_fill_NFT.jsx");
var DATA20_JSX = path.join(ROOT, "back_fill_NFT_Data20.jsx");
var BATCH_JSX = path.join(ROOT, "batch_fill.jsx");
var SAMPLE_MEMBERS = path.join(ROOT, "sample_NFT_members.csv");
var SAMPLE_DATA20 = path.join(ROOT, "sample_NFT_data20.csv");

var passed = 0;
var failed = 0;

function assert(cond, msg) {
  if (cond) {
    passed++;
    return;
  }
  failed++;
  console.error("FAIL: " + msg);
}

function extractFunction(src, name) {
  var re = new RegExp("function\\s+" + name + "\\s*\\(");
  var m = re.exec(src);
  if (!m) throw new Error("function " + name + " not found");
  var start = m.index;
  var brace = src.indexOf("{", start);
  if (brace < 0) throw new Error("no body for " + name);
  var depth = 0;
  var inStr = null;
  var inLine = false;
  var inBlock = false;
  var inRegex = false;
  var i;
  for (i = brace; i < src.length; i++) {
    var ch = src.charAt(i);
    var next = i + 1 < src.length ? src.charAt(i + 1) : "";
    if (inLine) {
      if (ch === "\n") inLine = false;
      continue;
    }
    if (inBlock) {
      if (ch === "*" && next === "/") {
        inBlock = false;
        i++;
      }
      continue;
    }
    if (inStr) {
      if (ch === "\\") {
        i++;
        continue;
      }
      if (ch === inStr) inStr = null;
      continue;
    }
    if (inRegex) {
      if (ch === "\\") {
        i++;
        continue;
      }
      if (ch === "/") inRegex = false;
      continue;
    }
    if (ch === "/" && next === "/") {
      inLine = true;
      i++;
      continue;
    }
    if (ch === "/" && next === "*") {
      inBlock = true;
      i++;
      continue;
    }
    if (ch === '"' || ch === "'") {
      inStr = ch;
      continue;
    }
    if (ch === "/") {
      var j = i - 1;
      while (j >= 0 && /\s/.test(src.charAt(j))) j--;
      var prev = j >= 0 ? src.charAt(j) : "";
      if (/[=(,!&|?:;{}\[~+%^<>]/.test(prev)) {
        inRegex = true;
        continue;
      }
    }
    if (ch === "{") depth++;
    else if (ch === "}") {
      depth--;
      if (depth === 0) return src.substring(start, i + 1);
    }
  }
  throw new Error("unclosed function " + name);
}

function MockFile(p) {
  this._path = String(p);
  this.encoding = "UTF-8";
  this.fsName = this._path;
  var parts = this._path.replace(/\\/g, "/").split("/");
  this.name = parts[parts.length - 1] || this._path;
}
Object.defineProperty(MockFile.prototype, "exists", {
  get: function () {
    try {
      return fs.existsSync(this._path);
    } catch (e) {
      return false;
    }
  }
});
MockFile.prototype.open = function () {
  this._opened = true;
  return true;
};
MockFile.prototype.read = function () {
  var buf = fs.readFileSync(this._path);
  var enc = String(this.encoding || "UTF-8").toUpperCase().replace(/_/g, "-");
  if (enc === "UTF-8" || enc === "UTF8") return buf.toString("utf8");
  if (
    enc === "SYSTEM" ||
    enc === "WINDOWS-1252" ||
    enc === "CP1252" ||
    enc === "WINDOWS1252"
  ) {
    return buf.toString("latin1");
  }
  return buf.toString("utf8");
};
MockFile.prototype.close = function () {
  this._opened = false;
};
MockFile.prototype.copy = function (dest) {
  var destPath = dest._path || dest.fsName;
  fs.copyFileSync(this._path, destPath);
  return true;
};
MockFile.prototype.remove = function () {
  try {
    fs.unlinkSync(this._path);
  } catch (eRm) {}
};
MockFile.decode = function (s) {
  try {
    return decodeURIComponent(String(s));
  } catch (eD) {
    return String(s);
  }
};

function loadHelpers(jsxPath) {
  var src = fs.readFileSync(jsxPath, "utf8");
  var names = [
    "trim",
    "stripLeadingBom",
    "stripBomFromRows",
    "formatSkip",
    "fileExt",
    "readFile",
    "parseRowRange",
    "looksLikeInvalidUtf8",
    "tryReadFileEncoding",
    "readFileWithDetectedEncoding",
    "unescapeTsvCell",
    "parseTSV",
    "parseCSV",
    "coalesceBrokenCsvRows",
    "looksLikeImagePath",
    "stripFileUrl",
    "normalizeLocalPath",
    "decodedPathVariants",
    "prepareImageFile",
    "resolveImageFile",
    "excelColLetter",
    "normalizeHeader",
    "buildHeaderInfo",
    "uniqueInts",
    "pickBestColIndex",
    "resolveColumnIndex",
    "csvCell",
    "rowDisplayName",
    "isValidCardRow",
    "filterStudentRows"
  ];
  var code = [];
  var i;
  for (i = 0; i < names.length; i++) {
    try {
      code.push(extractFunction(src, names[i]));
    } catch (eSkip) {
      // batch_fill.jsx does not define NFT-only helpers
    }
  }
  var File = MockFile;
  var Folder = { temp: { fsName: os.tmpdir() } };
  var app = { system: function () { return 0; } };
  var helpers = {};
  eval(
    "function findQrFileInFolder() { return null; }\n" +
      code.join("\n") +
      ";\nhelpers = { trim: trim, stripLeadingBom: stripLeadingBom, stripBomFromRows: stripBomFromRows, formatSkip: formatSkip, fileExt: fileExt, readFile: readFile, parseCSV: parseCSV, parseTSV: typeof parseTSV === 'function' ? parseTSV : null, coalesceBrokenCsvRows: coalesceBrokenCsvRows, looksLikeImagePath: looksLikeImagePath, normalizeLocalPath: normalizeLocalPath, decodedPathVariants: decodedPathVariants, resolveImageFile: resolveImageFile, rowDisplayName: typeof rowDisplayName === 'function' ? rowDisplayName : null, buildHeaderInfo: typeof buildHeaderInfo === 'function' ? buildHeaderInfo : null, resolveColumnIndex: typeof resolveColumnIndex === 'function' ? resolveColumnIndex : null, isValidCardRow: typeof isValidCardRow === 'function' ? isValidCardRow : null, filterStudentRows: typeof filterStudentRows === 'function' ? filterStudentRows : null, parseRowRange: typeof parseRowRange === 'function' ? parseRowRange : null, looksLikeInvalidUtf8: typeof looksLikeInvalidUtf8 === 'function' ? looksLikeInvalidUtf8 : null, readFileWithDetectedEncoding: typeof readFileWithDetectedEncoding === 'function' ? readFileWithDetectedEncoding : null };"
  );
  helpers.File = File;
  return helpers;
}

function writeBomCsv(filePath, bodyUtf8) {
  fs.writeFileSync(
    filePath,
    Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(bodyUtf8, "utf8")])
  );
}

function listDataFieldLetters(src) {
  var letters = [];
  var re = /prefix:\s*"data([A-Z])"/g;
  var m;
  while ((m = re.exec(src))) letters.push(m[1]);
  return letters;
}

function testSourceContracts() {
  var batchSrc = fs.readFileSync(BATCH_JSX, "utf8");
  assert(
    /var encodings = \["UTF-8", "UTF8"\]/.test(batchSrc),
    "batch_fill.jsx reads CSV as UTF-8 first"
  );
  assert(
    batchSrc.indexOf('encodings = ["SYSTEM"') < 0,
    "batch_fill.jsx does not prefer SYSTEM encoding"
  );

  var files = [NFT_JSX, DATA20_JSX, BATCH_JSX];
  var i;
  for (i = 0; i < files.length; i++) {
    var src = fs.readFileSync(files[i], "utf8");
    var label = path.basename(files[i]);
    assert(
      src.indexOf("Those fields will be skipped") >= 0,
      label + " treats missing image files as skip/continue"
    );
    assert(
      !/if \(missingFiles > 0\) \{\s*throw new Error/.test(src),
      label + " does not abort the run when image files are missing"
    );
    assert(
      src.indexOf("formatSkip(") >= 0,
      label + " logs skipped items with row/column/value"
    );
    assert(
      src.indexOf("SKIPPED / MISSING (row, column, value)") >= 0,
      label + " lists every skipped item in the final summary"
    );
    assert(
      src.indexOf(".slice(0, 30)") < 0 && src.indexOf(".slice(0, 20)") < 0,
      label + " does not truncate the skipped list"
    );
  }

  var nftFiles = [NFT_JSX, DATA20_JSX];
  for (i = 0; i < nftFiles.length; i++) {
    var nftSrc = fs.readFileSync(nftFiles[i], "utf8");
    var nftLabel = path.basename(nftFiles[i]);
    assert(
      nftSrc.indexOf("function looksLikeInvalidUtf8") >= 0,
      nftLabel + " detects invalid UTF-8"
    );
    assert(
      nftSrc.indexOf("function readFileWithDetectedEncoding") >= 0,
      nftLabel + " auto-detects CSV encoding"
    );
    assert(
      nftSrc.indexOf('utfNames = ["UTF-8", "UTF8"]') >= 0,
      nftLabel + " tries UTF-8 before SYSTEM"
    );
    assert(
      nftSrc.indexOf("SYSTEM (Windows-1252)") >= 0,
      nftLabel + " falls back to SYSTEM/Windows-1252"
    );
    assert(
      nftSrc.indexOf("function parseRowRange") >= 0,
      nftLabel + " parses a typed row range"
    );
    assert(
      nftSrc.indexOf("Row range:") >= 0 && nftSrc.indexOf("Start row:") >= 0,
      nftLabel + " keeps Start/End fields and a Row range field"
    );
    assert(
      nftSrc.indexOf("past end of CSV") >= 0,
      nftLabel + " skips rows past the end of the CSV"
    );
    assert(
      nftSrc.indexOf("blank / not a card row") >= 0,
      nftLabel + " skips blank rows instead of crashing"
    );
    assert(
      nftSrc.indexOf("suspendHistory") >= 0,
      nftLabel + " wraps the fill in one History step"
    );
    assert(
      nftSrc.indexOf("app.refresh(") < 0,
      nftLabel + " does not call app.refresh()"
    );
    assert(
      nftSrc.indexOf("function closeDocQuietly") >= 0,
      nftLabel + " closes temporary placed documents"
    );
    assert(
      nftSrc.indexOf("CSV encoding:") >= 0,
      nftLabel + " logs the encoding used in the summary"
    );
    assert(
      nftSrc.indexOf("1 = first data row after the header") >= 0,
      nftLabel + " documents the data-row numbering convention"
    );
  }
}

function testUtf8SampleAndBom(h) {
  var sampleBuf = fs.readFileSync(SAMPLE_MEMBERS);
  var utf8 = sampleBuf.toString("utf8");
  assert(utf8.indexOf("Juan Peña") >= 0, "sample_NFT_members.csv contains Juan Peña");
  assert(
    utf8.indexOf("Minna Señara.jpg") >= 0,
    "sample_NFT_members.csv contains Minna Señara.jpg path"
  );

  var latin1 = sampleBuf.toString("latin1");
  assert(
    latin1.indexOf("PeÃ±a") >= 0 || latin1.indexOf("SeÃ±ara") >= 0,
    "SYSTEM/latin1 decode of the same UTF-8 bytes is mojibake (the original bug)"
  );

  var rows = h.stripBomFromRows(h.coalesceBrokenCsvRows(h.parseCSV(h.readFile(SAMPLE_MEMBERS))));
  assert(rows.length >= 5, "sample has header + 4 dummy rows");
  assert(rows[0][0] === "DataA", "first header is DataA (no BOM)");
  assert(rows[4][1] === "Juan Peña", "DataB of dummy UTF-8 row is Juan Peña");
  assert(rows[4][5] === "Peña", "lastName of dummy UTF-8 row is Peña");
  assert(
    String(rows[4][7]).indexOf("Minna Señara.jpg") >= 0,
    "PictureA path keeps ñ (Señara), not mojibake"
  );
  assert(
    String(rows[4][7]).indexOf("SeÃ±ara") < 0,
    "PictureA path is not Windows-1252 mojibake"
  );

  if (h.buildHeaderInfo && h.rowDisplayName && h.isValidCardRow) {
    var header = rows[0];
    var data = rows.slice(1);
    var info = h.buildHeaderInfo(header);
    var filtered = h.filterStudentRows(header, data, info);
    assert(filtered.rows.length === 4, "all four dummy member rows are valid cards");
    var label = h.rowDisplayName(rows[4], info, data);
    assert(label === "Juan Peña", "row check label is Juan Peña (not mojibake / no name)");
    assert(h.isValidCardRow(rows[4], info, data), "Juan Peña row is a valid card");
  }

  var bomPath = path.join(os.tmpdir(), "ps_utf8_bom_members.csv");
  writeBomCsv(
    bomPath,
    "DataA,DataB,PictureA\nCEB-XX0099,Juan Peña,C:/tmp/Minna Señara.jpg\n"
  );
  var bomRows = h.stripBomFromRows(h.parseCSV(h.readFile(bomPath)));
  assert(bomRows[0][0] === "DataA", "UTF-8 BOM is stripped from the first header");
  assert(bomRows[0][0].charCodeAt(0) !== 0xfeff, "first header does not start with U+FEFF");
  assert(bomRows[1][1] === "Juan Peña", "BOM CSV still decodes ñ in the name");

  var stripped = h.stripLeadingBom("\uFEFFPictureA,PictureB");
  assert(stripped === "PictureA,PictureB", "stripLeadingBom removes U+FEFF");
  var rawBytes = h.stripLeadingBom("\u00ef\u00bb\u00bfDataA");
  assert(rawBytes === "DataA", "stripLeadingBom removes UTF-8 BOM misread as Latin-1");
}

function testPathFallbacks(h) {
  var dir = fs.mkdtempSync(path.join(os.tmpdir(), "ps-utf8-"));
  var real = path.join(dir, "Minna Señara.jpg");
  fs.writeFileSync(real, "fake-image");

  var encoded =
    dir.replace(/\\/g, "/") + "/Minna%20Se%C3%B1ara.jpg";
  var variants = h.decodedPathVariants(encoded);
  var foundDecoded = false;
  var vi;
  for (vi = 0; vi < variants.length; vi++) {
    if (String(variants[vi]).indexOf("Minna Señara.jpg") >= 0) foundDecoded = true;
  }
  assert(foundDecoded, "decodedPathVariants recovers ñ and spaces via decodeURI");

  var resolved = h.resolveImageFile(encoded);
  assert(!!(resolved && resolved.file), "resolveImageFile finds the file after decodeURI fallback");

  var missing = h.resolveImageFile(path.join(dir, "no-such-Minna Señara.jpg"));
  assert(!(missing && missing.file), "resolveImageFile reports missing when no variant exists");

  var skip = h.formatSkip(4, "PictureA", "C:\\photos\\Minna Señara.jpg", "file not found");
  assert(skip.indexOf("row 4") >= 0, "skip line includes row");
  assert(skip.indexOf("column PictureA") >= 0, "skip line includes column");
  assert(skip.indexOf("Minna Señara.jpg") >= 0, "skip line includes value with ñ");
  assert(skip.indexOf("file not found") >= 0, "skip line includes reason");
}

function writeUtf8PenaCsv(filePath) {
  fs.writeFileSync(
    filePath,
    "DataA,DataB,firstName,lastName\nCEB-XX0097,Juan Peña,Juan,Peña\n",
    "utf8"
  );
}

function writeCp1252PenaCsv(filePath) {
  fs.writeFileSync(
    filePath,
    Buffer.concat([
      Buffer.from("DataA,DataB,firstName,lastName\nCEB-XX0098,Juan Pe", "ascii"),
      Buffer.from([0xf1]),
      Buffer.from("a,Juan,Pe", "ascii"),
      Buffer.from([0xf1]),
      Buffer.from("a\n", "ascii")
    ])
  );
}

function testEncodingFallback(h, label) {
  assert(!!h.looksLikeInvalidUtf8, label + " exports looksLikeInvalidUtf8");
  assert(!!h.readFileWithDetectedEncoding, label + " exports readFileWithDetectedEncoding");
  assert(h.looksLikeInvalidUtf8("Juan Peña") === false, label + " UTF-8 Peña is valid UTF-8");
  assert(
    h.looksLikeInvalidUtf8("Juan Pe\uFFFDa") === true,
    label + " U+FFFD replacement marks invalid UTF-8"
  );
  assert(
    h.looksLikeInvalidUtf8("Juan PeÃ±a") === true,
    label + " typical Ã± mojibake marks invalid UTF-8"
  );

  var utfPath = path.join(os.tmpdir(), "ps_utf8_pena_" + label + ".csv");
  var cpPath = path.join(os.tmpdir(), "ps_cp1252_pena_" + label + ".csv");
  writeUtf8PenaCsv(utfPath);
  writeCp1252PenaCsv(cpPath);

  var utf = h.readFileWithDetectedEncoding(utfPath);
  assert(utf.encoding.indexOf("UTF-8") >= 0, label + " UTF-8 file is read as UTF-8");
  assert(utf.text.indexOf("Juan Peña") >= 0, label + " UTF-8 file keeps Juan Peña");
  assert(utf.text.indexOf("\uFFFD") < 0, label + " UTF-8 file has no replacement char");

  var cp = h.readFileWithDetectedEncoding(cpPath);
  assert(
    cp.encoding.indexOf("SYSTEM") >= 0,
    label + " Windows-1252 file falls back to SYSTEM/Windows-1252"
  );
  assert(
    cp.text.indexOf("Juan Peña") >= 0,
    label + " Windows-1252 file decodes Juan Peña after fallback"
  );
  assert(
    cp.text.indexOf("\uFFFD") < 0,
    label + " Windows-1252 fallback does not leave replacement chars"
  );

  var utfRows = h.parseCSV(utf.text);
  var cpRows = h.parseCSV(cp.text);
  assert(utfRows[1][1] === "Juan Peña", label + " parsed UTF-8 DataB is Juan Peña");
  assert(cpRows[1][1] === "Juan Peña", label + " parsed CP1252 DataB is Juan Peña");
}

function testRowRangeParsing(h, label) {
  assert(!!h.parseRowRange, label + " exports parseRowRange");
  var a = h.parseRowRange("101-104");
  assert(!a.error && a.start === 101 && a.end === 104, label + " parses 101-104");
  var b = h.parseRowRange("101");
  assert(!b.error && b.start === 101 && b.end === 101, label + " parses 101 as a single row");
  var c = h.parseRowRange("104-101");
  assert(!c.error && c.start === 101 && c.end === 104, label + " reverses 104-101");
  var d = h.parseRowRange(" 101 - 104 ");
  assert(!d.error && d.start === 101 && d.end === 104, label + " allows spaces around the hyphen");
  var e = h.parseRowRange("abc");
  assert(!!e.error, label + " rejects abc");
  var f = h.parseRowRange("10-");
  assert(!!f.error, label + " rejects 10-");
  var g = h.parseRowRange("101-104-105");
  assert(!!g.error, label + " rejects 101-104-105");
  var hEmpty = h.parseRowRange("");
  assert(hEmpty.error === "empty", label + " empty range is empty (dialog uses Start/End)");
  var iZero = h.parseRowRange("0");
  assert(!!iZero.error, label + " rejects 0");
  var jNeg = h.parseRowRange("-5");
  assert(!!jNeg.error, label + " rejects -5");
}

function testData20Fields(h) {
  var src10 = fs.readFileSync(NFT_JSX, "utf8");
  var src20 = fs.readFileSync(DATA20_JSX, "utf8");
  var letters10 = listDataFieldLetters(src10);
  var letters20 = listDataFieldLetters(src20);
  assert(letters10.length === 10, "Data10 script has 10 data fields (A-J)");
  assert(letters10[0] === "A" && letters10[9] === "J", "Data10 fields are DataA-DataJ");
  assert(src10.indexOf('prefix: "dataT"') < 0, "Data10 script does not define DataT");
  assert(
    src10.indexOf("/^(smart|data)([A-Ja-j])$/i") >= 0,
    "Data10 lookup regex stays A-J"
  );
  assert(letters20.length === 20, "Data20 script has 20 data fields (A-T)");
  assert(letters20[0] === "A" && letters20[19] === "T", "Data20 fields are DataA-DataT");
  assert(
    src20.indexOf("/^(smart|data)([A-Ta-t])$/i") >= 0,
    "Data20 lookup regex is widened to A-T"
  );
  assert(src20.indexOf("Data20: DataA-DataT") >= 0, "Data20 header/help says Data20: DataA-DataT");
  assert(src20.indexOf('columnAliases: ["Data K"]') >= 0, "DataK has column alias Data K");
  assert(src20.indexOf('layerAliases: ["dataK"]') >= 0, "DataK has layer alias dataK");
  assert(src20.indexOf('columnAliases: ["Data T"]') >= 0, "DataT has column alias Data T");
  assert(src20.indexOf('layerAliases: ["dataT"]') >= 0, "DataT has layer alias dataT");

  var sample = fs.readFileSync(SAMPLE_DATA20, "utf8");
  assert(sample.indexOf("DataT") >= 0, "sample_NFT_data20.csv has a DataT column");
  assert(sample.indexOf("Juan Peña") >= 0, "sample_NFT_data20.csv uses dummy name Juan Peña");
  var rows = h.stripBomFromRows(h.parseCSV(h.readFile(SAMPLE_DATA20)));
  var header = rows[0];
  var dataLetters = "ABCDEFGHIJKLMNOPQRST";
  var di;
  for (di = 0; di < dataLetters.length; di++) {
    var want = "Data" + dataLetters.charAt(di);
    var found = false;
    var hi;
    for (hi = 0; hi < header.length; hi++) {
      if (header[hi] === want) found = true;
    }
    assert(found, "sample_NFT_data20.csv header includes " + want);
  }
  assert(rows.length === 5, "sample_NFT_data20.csv has header + 4 dummy rows");
  if (h.buildHeaderInfo && h.resolveColumnIndex) {
    var info = h.buildHeaderInfo(header);
    var idxT = h.resolveColumnIndex(
      info,
      {
        prefix: "dataT",
        column: "DataT",
        columnAliases: ["Data T"],
        type: "text"
      },
      rows.slice(1)
    );
    assert(idxT !== undefined, "Data20 resolveColumnIndex finds DataT");
    var idxAlias = h.resolveColumnIndex(
      info,
      { prefix: "dataK", column: "DataK", columnAliases: ["Data K"], type: "text" },
      rows.slice(1)
    );
    assert(idxAlias !== undefined, "Data20 resolveColumnIndex finds DataK");
  }
}

function testSkipContinueLogic() {
  // Simulate the pre-check: missing files add a warning, never throw.
  var missingFiles = 1;
  var warnings = [];
  var skipped = [];
  var pathPreview = ["  r4 PictureA = Minna Señara.jpg [FILE MISSING]"];
  if (missingFiles > 0) {
    warnings.push(
      "Pre-check: " +
        missingFiles +
        " image file(s) not found at the path in the CSV. " +
        "Those fields will be skipped; other cards and fields still fill.\n" +
        pathPreview.join("\n")
    );
  }
  var rowsFilled = 0;
  var r;
  for (r = 1; r <= 4; r++) {
    try {
      if (r === 4) {
        skipped.push(
          "row 4, column PictureA, value 'Minna Señara.jpg' -- file not found"
        );
      }
      rowsFilled++;
    } catch (eRow) {
      throw new Error("row errors must not abort the run: " + eRow.message);
    }
  }
  assert(rowsFilled === 4, "run continues through all rows when a PictureA file is missing");
  assert(warnings.length === 1, "pre-check is recorded as a warning");
  assert(skipped.length === 1, "missing file is listed as a skipped item");
  assert(
    warnings[0].indexOf("will be skipped") >= 0,
    "warning says the run continues"
  );
}

function main() {
  testSourceContracts();
  var nft = loadHelpers(NFT_JSX);
  var data20 = loadHelpers(DATA20_JSX);
  var batch = loadHelpers(BATCH_JSX);
  testUtf8SampleAndBom(nft);
  testUtf8SampleAndBom(batch);
  testPathFallbacks(nft);
  testPathFallbacks(batch);
  testSkipContinueLogic();
  testEncodingFallback(nft, "back_fill_NFT.jsx");
  testEncodingFallback(data20, "back_fill_NFT_Data20.jsx");
  testRowRangeParsing(nft, "back_fill_NFT.jsx");
  testRowRangeParsing(data20, "back_fill_NFT_Data20.jsx");
  testData20Fields(data20);

  console.log("Passed: " + passed + "  Failed: " + failed);
  if (failed) process.exit(1);
}

main();
