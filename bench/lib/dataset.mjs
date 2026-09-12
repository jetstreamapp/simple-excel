/**
 * Deterministic, seeded dataset generators for the xlsx engine benchmark.
 *
 * Every dataset is defined as a header (column names) plus a per-row factory driven by a seeded
 * xorshift PRNG, so the same `{ dataset, rows, seed }` always yields byte-identical rows regardless
 * of which engine consumes them. Rows are produced lazily by a generator (`createRowIterator`);
 * nothing is materialised unless the caller asks for it (`materialize`).
 *
 * Pure JS with no Node built-ins so the same module can be bundled for the browser bench.
 */

export const DEFAULT_SEED = 20260911;

/** Excel's hard limit on characters per cell. Oversized cells here are intentional (see `mixed`). */
export const EXCEL_MAX_CELL_CHARS = 32_767;

const ALNUM = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
const UPPER = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

const FIRST_NAMES = [
  'Aiko',
  'Björn',
  'Chloé',
  'Dmitri',
  'Emeka',
  'Fátima',
  'Gunnar',
  'Hélène',
  'Ísak',
  'José',
  'Kenji',
  'Léa',
  'María',
  'Nikolai',
  'Øyvind',
  'Priya',
  'Quentin',
  'Renée',
  'Søren',
  'Tomás',
  'Ülkü',
  'Valérie',
  'Wei',
  'Xóchitl',
  'Yusuf',
  'Zoë',
  '太郎',
  '美咲',
  '민준',
  '서연',
  'Владимир',
  'Ольга',
  'محمد',
  'فاطمة',
  'Ελένη',
  'Γιώργος',
];
const LAST_NAMES = [
  'Andersson',
  'Bäcker',
  'Çelik',
  'Dubois',
  'Eriksen',
  'Fernández',
  'García',
  'Håkansson',
  'Ibáñez',
  'Jørgensen',
  'Kowalski',
  'Løken',
  'Müller',
  'Nguyễn',
  "O'Brien",
  'Pérez',
  'Quiñones',
  'Rousseau',
  'Schäfer',
  'Takahashi',
  'Ünal',
  'Villanueva',
  'Wójcik',
  'Xu',
  'Yamamoto',
  'Zhāng',
  '佐藤',
  '鈴木',
  '김',
  '이',
  'Иванов',
  'Смирнова',
];
const INDUSTRIES = [
  'Agriculture',
  'Banking',
  'Biotechnology',
  'Communications',
  'Construction',
  'Education',
  'Energy',
  'Finance',
  'Healthcare',
  'Manufacturing',
  'Media',
  'Retail',
  'Technology',
  'Transportation',
  'Utilities',
];
const ACCOUNT_TYPES = ['Customer - Direct', 'Customer - Channel', 'Prospect', 'Partner', 'Installation Partner', 'Other'];
const WORDS = [
  'account',
  'pipeline',
  'quarter',
  'renewal',
  'discount',
  'escalation',
  'territory',
  'forecast',
  'opportunity',
  'contract',
  'invoice',
  'support',
  'onboarding',
  'migration',
  'integration',
  'dashboard',
  'workflow',
  'approval',
  'campaign',
  'segment',
  'naïve',
  'résumé',
  'façade',
  'coöperate',
  '日本語',
  'データ',
  'привет',
  'مرحبا',
  '🚀',
  '✅',
];

/**
 * xorshift32 PRNG. Small, fast, deterministic — quality is irrelevant here, reproducibility is not.
 */
export function createRng(seed = DEFAULT_SEED) {
  let state = seed >>> 0 || 0x9e3779b9;
  const next = () => {
    state ^= state << 13;
    state >>>= 0;
    state ^= state >>> 17;
    state ^= state << 5;
    state >>>= 0;
    return state / 0x1_0000_0000;
  };
  const int = (min, max) => min + Math.floor(next() * (max - min + 1));
  const pick = list => list[Math.floor(next() * list.length)];
  const chance = probability => next() < probability;
  const chars = (alphabet, length) => {
    let out = '';
    for (let i = 0; i < length; i++) {
      out += alphabet[Math.floor(next() * alphabet.length)];
    }
    return out;
  };
  return { next, int, pick, chance, chars };
}

const salesforceId = (rng, prefix) => prefix + rng.chars(ALNUM, 12) + rng.chars(UPPER, 3);
const personName = rng => `${rng.pick(FIRST_NAMES)} ${rng.pick(LAST_NAMES)}`;
const sentence = rng => {
  const wordCount = rng.int(4, 14);
  const words = [];
  for (let i = 0; i < wordCount; i++) {
    words.push(rng.pick(WORDS));
  }
  return words.join(' ') + '.';
};
const paragraph = (rng, sentenceCount) => {
  const sentences = [];
  for (let i = 0; i < sentenceCount; i++) {
    sentences.push(sentence(rng));
  }
  return sentences.join('\n');
};
const dateTime = rng =>
  new Date(Date.UTC(2015 + rng.int(0, 11), rng.int(0, 11), rng.int(1, 28), rng.int(0, 23), rng.int(0, 59), rng.int(0, 59)));
const dateOnly = rng => new Date(Date.UTC(2015 + rng.int(0, 11), rng.int(0, 11), rng.int(1, 28)));
const decimal = (rng, max, places) => Number((rng.next() * max).toFixed(places));

/** JSON-stringified subquery result the way Jetstream flattens `Contacts__r` into a single cell. */
const subqueryBlob = (rng, oversized) => {
  const recordCount = oversized ? 260 : rng.int(1, 4);
  const records = [];
  for (let i = 0; i < recordCount; i++) {
    records.push({
      attributes: { type: 'Contact', url: `/services/data/v62.0/sobjects/Contact/${salesforceId(rng, '003')}` },
      Id: salesforceId(rng, '003'),
      Name: personName(rng),
      Email: `${rng.chars(ALNUM, 8).toLowerCase()}@example.com`,
      Title: sentence(rng),
    });
  }
  return JSON.stringify({ totalSize: recordCount, done: true, records });
};

const NULL_RATE = 0.1;
const OVERSIZED_RATE = 0.01;

const DATASETS = {
  mixed: {
    description:
      'Salesforce-shaped export: 18-char Ids, unicode names, integers and decimals, booleans, JS Date objects, ISO datetime strings, ' +
      'long text with newlines, JSON subquery blobs (1% of them over 32,767 chars) and 10% nulls in every column except Id.',
    columns: [
      'Id',
      'Name',
      'AccountNumber',
      'Amount',
      'Quantity',
      'Probability',
      'IsActive',
      'IsDeleted',
      'CreatedDate',
      'LastModifiedDate',
      'CloseDate',
      'Description',
      'Contacts__r',
      'Owner.Name',
      'Owner.Id',
      'Type',
      'Industry',
      'Website',
      'AnnualRevenue',
      'NumberOfEmployees',
    ],
    row(rng, rowIndex) {
      const nullable = produce => (rng.chance(NULL_RATE) ? null : produce());
      return [
        salesforceId(rng, '001'),
        nullable(() => personName(rng)),
        nullable(() => `ACC-${String(rowIndex + 1).padStart(8, '0')}`),
        nullable(() => decimal(rng, 1_000_000, 2)),
        nullable(() => rng.int(0, 10_000)),
        nullable(() => decimal(rng, 100, 1)),
        nullable(() => rng.chance(0.5)),
        nullable(() => rng.chance(0.05)),
        nullable(() => dateTime(rng)),
        nullable(() => dateTime(rng).toISOString()),
        nullable(() => dateOnly(rng)),
        nullable(() => paragraph(rng, rng.int(2, 8))),
        nullable(() => subqueryBlob(rng, rng.chance(OVERSIZED_RATE))),
        nullable(() => personName(rng)),
        nullable(() => salesforceId(rng, '005')),
        nullable(() => rng.pick(ACCOUNT_TYPES)),
        nullable(() => rng.pick(INDUSTRIES)),
        nullable(() => `https://www.${rng.chars(ALNUM, 10).toLowerCase()}.example.com`),
        nullable(() => rng.int(0, 5_000_000_000)),
        nullable(() => rng.int(1, 250_000)),
      ];
    },
  },

  wide: {
    description:
      '107 columns, no nulls: 60 strings (random words, 5–30 chars), 30 numbers (half integers, half decimals), 10 JS Date objects, 7 booleans.',
    columns: [
      ...Array.from({ length: 60 }, (_, i) => `Text_${String(i + 1).padStart(2, '0')}`),
      ...Array.from({ length: 30 }, (_, i) => `Num_${String(i + 1).padStart(2, '0')}`),
      ...Array.from({ length: 10 }, (_, i) => `Date_${String(i + 1).padStart(2, '0')}`),
      ...Array.from({ length: 7 }, (_, i) => `Flag_${String(i + 1).padStart(2, '0')}`),
    ],
    row(rng) {
      const cells = Array.from({ length: 107 });
      for (let i = 0; i < 60; i++) {
        cells[i] = rng.chars(ALNUM, rng.int(5, 30));
      }
      for (let i = 60; i < 90; i++) {
        cells[i] = i % 2 === 0 ? rng.int(0, 1_000_000) : decimal(rng, 100_000, 4);
      }
      for (let i = 90; i < 100; i++) {
        cells[i] = dateTime(rng);
      }
      for (let i = 100; i < 107; i++) {
        cells[i] = rng.chance(0.5);
      }
      return cells;
    },
  },

  'strings-unique': {
    description:
      '20 columns of globally unique 24-char strings (row/column prefix + random tail) — the shared-string-table worst case: nothing dedupes.',
    columns: Array.from({ length: 20 }, (_, i) => `Key_${String(i + 1).padStart(2, '0')}`),
    row(rng, rowIndex) {
      const rowPrefix = rowIndex.toString(36).padStart(6, '0');
      const cells = Array.from({ length: 20 });
      for (let i = 0; i < 20; i++) {
        cells[i] = `${rowPrefix}${i.toString(36)}${rng.chars(ALNUM, 17)}`;
      }
      return cells;
    },
  },

  numeric: {
    description: '20 numeric columns, no nulls: 10 integers (0–1e6) and 10 decimals (4 places).',
    columns: Array.from({ length: 20 }, (_, i) =>
      i < 10 ? `Int_${String(i + 1).padStart(2, '0')}` : `Dec_${String(i - 9).padStart(2, '0')}`,
    ),
    row(rng) {
      const cells = Array.from({ length: 20 });
      for (let i = 0; i < 10; i++) {
        cells[i] = rng.int(0, 1_000_000);
      }
      for (let i = 10; i < 20; i++) {
        cells[i] = decimal(rng, 1_000_000, 4);
      }
      return cells;
    },
  },
};

export const DATASET_NAMES = Object.keys(DATASETS);

/** Plain row-count sizes. */
export const SIZES = {
  '1k': 1_000,
  '10k': 10_000,
  '100k': 100_000,
  '250k': 250_000,
  '500k': 500_000,
  '1m': 1_000_000,
};

/**
 * Preset sizes that pin the dataset as well as the row count. `18m-cells` reproduces the historical
 * SheetJS "Invalid array length" failure shape (900,000 rows × 20 mixed columns).
 */
export const PRESET_SIZES = {
  '18m-cells': { dataset: 'mixed', rows: 900_000 },
  'wide-100k': { dataset: 'wide', rows: 100_000 },
};

export const SIZE_NAMES = [...Object.keys(SIZES), ...Object.keys(PRESET_SIZES)];

export function getDataset(name) {
  const dataset = DATASETS[name];
  if (!dataset) {
    throw new Error(`Unknown dataset "${name}". Known: ${DATASET_NAMES.join(', ')}`);
  }
  return { name, description: dataset.description, columns: dataset.columns, columnCount: dataset.columns.length };
}

/**
 * Resolve `(dataset, size)` into a concrete cell shape. Preset sizes override the dataset, so
 * `--datasets numeric --sizes 18m-cells` still means mixed × 900k (callers dedupe the result).
 */
export function resolveShape(datasetName, sizeName) {
  const preset = PRESET_SIZES[sizeName];
  if (preset) {
    const dataset = getDataset(preset.dataset);
    return { dataset: dataset.name, size: sizeName, rows: preset.rows, columns: dataset.columns, cells: preset.rows * dataset.columnCount };
  }
  const rows = SIZES[sizeName];
  if (!rows) {
    throw new Error(`Unknown size "${sizeName}". Known: ${SIZE_NAMES.join(', ')}`);
  }
  const dataset = getDataset(datasetName);
  return { dataset: dataset.name, size: sizeName, rows, columns: dataset.columns, cells: rows * dataset.columnCount };
}

/**
 * Lazily yields `rows` data rows (header excluded) for the dataset. Each call creates a fresh PRNG,
 * so iterating twice with the same arguments yields identical rows.
 */
export function* createRowIterator({ dataset, rows, seed = DEFAULT_SEED }) {
  const definition = DATASETS[dataset];
  if (!definition) {
    throw new Error(`Unknown dataset "${dataset}". Known: ${DATASET_NAMES.join(', ')}`);
  }
  const rng = createRng(seed);
  for (let rowIndex = 0; rowIndex < rows; rowIndex++) {
    yield definition.row(rng, rowIndex);
  }
}

/** Materialise every row into an array — what SheetJS (and Jetstream's in-memory query results) need. */
export function materialize(options) {
  const out = Array.from({ length: options.rows });
  let i = 0;
  for (const row of createRowIterator(options)) {
    out[i++] = row;
  }
  return out;
}
