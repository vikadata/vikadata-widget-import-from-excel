const assert = require("node:assert/strict");
const { readFileSync } = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");
const ts = require("typescript");
const XLSX = require("xlsx");

function loadSource(filename, imports = {}, globals = {}) {
  const sourcePath = path.resolve(__dirname, "..", filename);
  const { outputText } = ts.transpileModule(readFileSync(sourcePath, "utf8"), {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2018,
      jsx: ts.JsxEmit.React,
      esModuleInterop: true,
    },
  });
  const module = { exports: {} };
  vm.runInNewContext(outputText, {
    module,
    exports: module.exports,
    require(name) {
      if (Object.hasOwn(imports, name)) return imports[name];
      if (name.startsWith(".")) {
        return loadSource(path.resolve(path.dirname(sourcePath), name + ".ts"));
      }
      return require(name);
    },
    ...globals,
  }, { filename: sourcePath });
  return module.exports;
}

function findInput(element) {
  if (element.type === "input") return element;
  for (const child of element.children || []) {
    if (child && typeof child === "object") {
      const found = findInput(child);
      if (found) return found;
    }
  }
}

const fields = [
  { id: "amount", name: "Amount", type: "Currency" },
  { id: "count", name: "Count", type: "Number" },
  { id: "account", name: "Account", type: "Text" },
  { id: "date", name: "Date", type: "DateTime" },
];

async function importSheet(rows, options = {}) {
  const worksheet = XLSX.utils.aoa_to_sheet(rows, { origin: options.origin || "A1" });
  if (options.origin) {
    worksheet["!ref"] = XLSX.utils.encode_range({
      s: XLSX.utils.decode_cell(options.origin),
      e: XLSX.utils.decode_range(worksheet["!ref"]).e,
    });
  }
  if (worksheet.A2 && worksheet.A2.t === "n") worksheet.A2.z = "#,##0.00";
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  const binary = XLSX.write(workbook, {
    type: "binary",
    bookType: options.bookType || "xlsx",
  });
  const writes = [];
  const alerts = [];
  const progress = [];
  const readers = [];
  const react = {
    createElement(type, props, ...children) { return { type, props, children }; },
    createRef() { return { current: { value: "fixture" } }; },
    useState(initial) { return [initial, (value) => progress.push(value)]; },
  };
  const { ExcelImport } = loadSource("src/excel_import.tsx", {
    react,
    "@vikadata/components": { Button: "button", Loading: "loading" },
    "@vikadata/widget-sdk": {
      useActiveViewId: () => "view",
      useFields: () => options.fields || fields,
      useDatasheet: () => ({
        addRecords(records) {
          writes.push(JSON.parse(JSON.stringify(records)));
          return Promise.resolve([]);
        },
      }),
    },
  }, {
    confirm: () => true,
    alert: (message) => alerts.push(String(message)),
    console: { log() {} },
    FileReader: class {
      constructor() { readers.push(this); }
      readAsBinaryString(binary) { this.result = binary; }
    },
    setTimeout() { throw new Error("Unexpected batch timer"); },
  });
  findInput(ExcelImport()).props.onChange({ target: { files: [binary] } });
  readers[0].onload({ target: { result: readers[0].result } });
  if (readers[0].onloadend) readers[0].onloadend();
  await Promise.resolve();
  return { writes, alerts, progress };
}

const screenshotAmounts = [
  [1200, 1200], [1145, 1145], [34300, 34300],
  ["28,800.00", 28800], ["6,730.00", 6730], [27550, 27550],
  ["100,000.00", 100000], ["6,500.00", 6500], ["5,800.00", 5800],
  ["11,000.00", 11000], ["22,000.00", 22000], ["11,950.00", 11950],
  ["12,400.00", 12400], ["65,090.00", 65090], ["97,500.00", 97500],
  ["7,525.00", 7525], ["80,000.00", 80000], ["30,000.00", 30000],
  ["11,500.00", 11500],
];

for (const bookType of ["xlsx", "biff8", "csv"]) {
  test(`${bookType}: imports mixed numeric cells and thousand-separated text`, async () => {
    const account = "001234567890123456789";
    const date = "2026/05/30 00:00:00";
    const result = await importSheet([
      ["Amount", "Count", "Account", "Date"],
      ...screenshotAmounts.map(([value]) => [value, value, account, date]),
    ], { bookType });
    assert.deepEqual(result.alerts, []);
    assert.equal(result.writes.length, 1);
    assert.deepEqual(result.writes[0].map(({ valuesMap }) => [valuesMap.amount, valuesMap.count]),
      screenshotAmounts.map(([, expected]) => [expected, expected]));
    if (bookType !== "csv") {
      assert.ok(result.writes[0].every(({ valuesMap }) => valuesMap.account === account));
      assert.ok(result.writes[0].every(({ valuesMap }) => valuesMap.date === new Date(date).getTime()));
    }
    assert.deepEqual(result.progress, [true, false]);
  });
}

test("preserves zero, negatives, decimals and blank numeric cells", async () => {
  const result = await importSheet([
    ["Amount", "Count", "Account"],
    [0, "0", "zero"],
    ["-1,200.50", "1.2e3", "negative"],
    ["$28,800.00", ".5", "currency"],
    ["", undefined, "blank"],
    ["   ", "\u00a0", "whitespace"],
  ]);
  assert.deepEqual(result.alerts, []);
  assert.deepEqual(result.writes[0].map(({ valuesMap }) => [valuesMap.amount, valuesMap.count]), [
    [0, 0], [-1200.5, 1200], [28800, 0.5], [null, null], [null, null],
  ]);
});

for (const [fieldName, fieldType] of [["Amount", "Currency"], ["Count", "Number"], ["Score", "Rating"]]) {
  test(`rejects invalid ${fieldType} values before writing any records`, async () => {
    const result = await importSheet([
      [fieldName], [1200], [], ["1,20.00"],
    ], { fields: [{ id: "value", name: fieldName, type: fieldType }] });
    assert.deepEqual(result.writes, []);
    assert.equal(result.alerts.length, 1);
    assert.match(result.alerts[0], /4/);
    assert.ok(result.alerts[0].includes(fieldName));
    assert.deepEqual(result.progress, [true, false]);
  });
}

test("reports actual sheet row and column even when the used range starts later", async () => {
  const result = await importSheet([
    ["Amount", "Count"], [1200, 1], [], ["not a number", 2],
  ], { origin: "C5" });
  assert.deepEqual(result.writes, []);
  assert.equal(result.alerts.length, 1);
  assert.match(result.alerts[0], /C8/);
  assert.ok(result.alerts[0].includes("Amount"));
});

test("does not validate computed fields or columns missing from the file", async () => {
  const result = await importSheet([
    ["Account", "Computed"], ["00123", "not a number"],
  ], { fields: [...fields, { id: "computed", name: "Computed", type: "Number", isComputed: true }] });
  assert.deepEqual(result.alerts, []);
  assert.deepEqual(result.writes, [[{ valuesMap: { account: "00123" } }]]);
});

test("maps incoming and outgoing amounts independently without shifting blank cells", async () => {
  const result = await importSheet([
    ["Outgoing", "Account", "Incoming"],
    ["21,000.00", "0001", 1200],
    [undefined, "0002", "28,800.00"],
    ["11,500.00", "0003"],
  ], { fields: [
    { id: "incoming", name: "Incoming", type: "Currency" },
    { id: "outgoing", name: "Outgoing", type: "Currency" },
    { id: "account", name: "Account", type: "Text" },
  ] });
  assert.deepEqual(result.alerts, []);
  assert.deepEqual(result.writes, [[
    { valuesMap: { incoming: 1200, outgoing: 21000, account: "0001" } },
    { valuesMap: { incoming: 28800, outgoing: null, account: "0002" } },
    { valuesMap: { incoming: null, outgoing: 11500, account: "0003" } },
  ]]);
});

test("validates all rows before scheduling any large-import batches", async () => {
  const result = await importSheet([
    ["Amount"], ...Array.from({ length: 2001 }, () => [1200]), ["1,20.00"],
  ]);
  assert.deepEqual(result.writes, []);
  assert.equal(result.alerts.length, 1);
  assert.match(result.alerts[0], /A2003/);
  assert.deepEqual(result.progress, [true, false]);
});

const { parseImportNumber } = loadSource("src/parse_number.ts");

test("parses common currency symbols, signs, grouped amounts and exponents", () => {
  for (const [input, expected] of [
    ...screenshotAmounts,
    ["$1,200.50", 1200.5], ["-$1,200.50", -1200.5], ["$-1,200.50", -1200.5],
    ["\u00a51,200.50", 1200.5], ["\uffe51\uff0c200.50", 1200.5],
    ["1,200.50\u20ac", 1200.5], ["\u00a31,200.50", 1200.5],
    ["(1,200.50)", -1200.5], ["($1,200.50)", -1200.5],
    ["1 200.50", 1200.5], ["1\u00a0200.50", 1200.5], ["1\u202f200.50", 1200.5],
    [" +1200.50 ", 1200.5], [".5", 0.5], ["-.5", -0.5], ["1200.", 1200],
    ["1.2e3", 1200], ["1.2E+3", 1200], ["-1.2e-3", -0.0012],
    [0, 0], ["0", 0], ["0.00", 0], [1.2e20, 1.2e20],
  ]) {
    assert.equal(parseImportNumber(input), expected, String(input));
  }
});

test("keeps empty numeric cells empty", () => {
  for (const input of [null, undefined, "", " ", "\t\n", "\u00a0"]) {
    assert.equal(parseImportNumber(input), null);
  }
});

test("rejects malformed or non-finite amounts rather than extracting partial digits", () => {
  for (const input of [
    "1,20.00", "12,34,567.00", "1 20.00", "1,200,00", "1.200,50", "1.2.3",
    "1e", "1e309", "NaN", "Infinity", "1200USD", "abc123", "123abc", "--",
    "(+1200)", "(-1200)", "1$200", true, {}, [], NaN, Infinity, -Infinity,
  ]) {
    assert.throws(() => parseImportNumber(input), /Invalid numeric value/, String(input));
  }
});
