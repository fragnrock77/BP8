const assert = require('assert');

const {
  tokenizeQuery,
  toPostfix,
  evaluateQuery,
  convertRowsToCsv,
  buildCaches,
  normalizeParsedData,
  extractKeywords,
 codex/audit-application-and-add-column-filtering-fi4qse
  getColumnsFor,
  getSingleFileColumns,

 codex/audit-application-and-add-column-filtering-yhgnz5
  getColumnsFor,
  getSingleFileColumns,

 codex/audit-application-and-add-column-filtering-09or3v
  getColumnsFor,
  getSingleFileColumns,

  getAvailableColumns,
 main
 main
 main
  __setTestState,
  __getTestState,
} = require('../app.js');

const results = [];

function test(name, fn) {
  try {
    fn();
    results.push({ name, status: 'passed' });
  } catch (error) {
    results.push({ name, status: 'failed', error });
  }
}

const SAMPLE_ROWS = [
  ['Alice', 'Premium', 'Active'],
  ['Bob', 'Standard', 'Inactive'],
  ['Charlie', 'Premium', 'Active'],
  ['Dora', 'Premium', 'Inactive'],
];

function resetStateForTests() {
  const columns = getSingleFileColumns(['Name', 'Plan', 'Status'], SAMPLE_ROWS);
  __setTestState({
    headers: ['Name', 'Plan', 'Status'],
    rawRows: SAMPLE_ROWS,
    filteredRows: SAMPLE_ROWS.slice(),
    rowTextCache: [],
    lowerRowTextCache: [],
    tableColumns: columns,
    matchesColumnIndex: -1,
    currentPage: 1,
    currentFileName: 'sample',
  });
  buildCaches();
}

resetStateForTests();

test('tokenizeQuery extracts operands, operators and quoted expressions', () => {
  const tokens = tokenizeQuery('premium AND "active user" OR (NOT standard)');
  const values = tokens.map((token) => token.value);
  assert.deepStrictEqual(values, [
    'premium',
    'AND',
    'active user',
    'OR',
    '(',
    'NOT',
    'standard',
    ')',
  ]);
});

resetStateForTests();

test('toPostfix throws on unbalanced parentheses', () => {
  const tokens = tokenizeQuery('(premium AND standard');
  assert.throws(() => toPostfix(tokens), /Parenthèses déséquilibrées/);
});

resetStateForTests();

test('toPostfix respects operator precedence', () => {
  const tokens = tokenizeQuery('premium AND NOT inactive OR standard');
  const postfix = toPostfix(tokens).map((token) => token.value);
  assert.deepStrictEqual(postfix, ['premium', 'inactive', 'NOT', 'AND', 'standard', 'OR']);
});

resetStateForTests();

test('evaluateQuery filters rows using boolean logic', () => {
  const options = { caseSensitive: false, exactMatch: false };
  const indexes = evaluateQuery(tokenizeQuery('premium AND NOT Bob'), options);
  assert.deepStrictEqual(indexes, [0, 2, 3]);
});

resetStateForTests();

test('evaluateQuery supports case sensitive and exact matches', () => {
  let indexes = evaluateQuery(tokenizeQuery('premium'), {
    caseSensitive: true,
    exactMatch: true,
  });
  assert.deepStrictEqual(indexes, []);

  indexes = evaluateQuery(tokenizeQuery('Active'), {
    caseSensitive: false,
    exactMatch: true,
  });
  assert.deepStrictEqual(indexes, [0, 2]);

  indexes = evaluateQuery(tokenizeQuery('Premium'), {
    caseSensitive: true,
    exactMatch: true,
  });
  assert.deepStrictEqual(indexes, [0, 2, 3]);
});

resetStateForTests();

test('convertRowsToCsv quotes separators and quotes', () => {
  const csv = convertRowsToCsv(['Name', 'Comment'], [
    ['Alice', 'simple'],
    ['Bob', 'needs, comma'],
    ['Charlie', 'He said "hello"'],
  ]);

  assert.strictEqual(
    csv,
    'Name,Comment\n' +
      'Alice,simple\n' +
      'Bob,"needs, comma"\n' +
      'Charlie,"He said ""hello"""'
  );
});

resetStateForTests();

test('buildCaches keeps caches synchronised', () => {
  const state = __getTestState();
  assert.strictEqual(state.rawRows.length, 4);
  assert.strictEqual(state.rowTextCache.length, 4);
  assert.ok(Array.isArray(state.rowTextCache[0]));
  assert.strictEqual(state.lowerRowTextCache[0][0], 'alice');
});

resetStateForTests();

test('normalizeParsedData infers headers and sanitises values', () => {
  const normalized = normalizeParsedData({
    headers: [],
    rows: [
      ['Nom', 'Âge'],
      ['Alice', 30],
      ['Bob', null],
    ],
  });
  assert.deepStrictEqual(normalized.headers, ['Nom', 'Âge']);
  assert.deepStrictEqual(normalized.rows, [
    ['Alice', '30'],
    ['Bob', ''],
  ]);
});

resetStateForTests();

test('extractKeywords returns unique trimmed entries', () => {
  const keywords = extractKeywords([
    ['  Alpha  ', ''],
    ['beta', 'Gamma'],
    ['beta', null],
  ]);
  assert.deepStrictEqual(keywords, ['Alpha', 'beta', 'Gamma']);
});

resetStateForTests();

 codex/audit-application-and-add-column-filtering-fi4qse

 codex/audit-application-and-add-column-filtering-yhgnz5

 codex/audit-application-and-add-column-filtering-09or3v
 main
 main
test('getColumnsFor prefixes labels for comparison datasets', () => {
  const columns = getColumnsFor('ref', [['A', 'B', 'C']], []);
  assert.deepStrictEqual(columns.map((col) => col.label), [
    'ref.Colonne 1',
    'ref.Colonne 2',
    'ref.Colonne 3',
  ]);
});

test('getSingleFileColumns keeps display labels without prefixes', () => {
  const columns = getSingleFileColumns(['Nom', 'Âge'], [
    ['Alice', '30'],
    ['Bob', '28'],
  ]);
  assert.deepStrictEqual(columns.map((col) => col.label), ['Nom', 'Âge']);
 codex/audit-application-and-add-column-filtering-fi4qse

 codex/audit-application-and-add-column-filtering-yhgnz5


test('getAvailableColumns falls back to generated labels', () => {
  __setTestState({
    headers: [],
    rawRows: [
      ['A', 'B', 'C'],
      ['D', 'E', 'F'],
    ],
  });
  const columns = getAvailableColumns([
    ['A', 'B', 'C'],
  ]);
  assert.deepStrictEqual(columns.map((col) => col.label), [
    'Colonne 1',
    'Colonne 2',
    'Colonne 3',
  ]);
 main
 main
 main
});

resetStateForTests();

test('evaluateQuery honours explicit column restriction', () => {
  const indexes = evaluateQuery(tokenizeQuery('premium'), {
    caseSensitive: false,
    exactMatch: false,
    columnIndexes: [1],
  });
  assert.deepStrictEqual(indexes, [0, 2, 3]);

  const nameIndexes = evaluateQuery(tokenizeQuery('alice'), {
    caseSensitive: false,
    exactMatch: false,
    columnIndexes: [0],
  });
  assert.deepStrictEqual(nameIndexes, [0]);

  const restrictedIndexes = evaluateQuery(tokenizeQuery('alice'), {
    caseSensitive: false,
    exactMatch: false,
    columnIndexes: [1],
  });
  assert.deepStrictEqual(restrictedIndexes, []);
});

const failed = results.filter((result) => result.status === 'failed');
results.forEach((result) => {
  if (result.status === 'passed') {
    console.log(`✔ ${result.name}`);
  } else {
    console.error(`✖ ${result.name}`);
    console.error(result.error);
  }
});

if (failed.length) {
  process.exitCode = 1;
  console.error(`\n${failed.length} test(s) failed.`);
} else {
  console.log(`\n${results.length} test(s) passed.`);
}
