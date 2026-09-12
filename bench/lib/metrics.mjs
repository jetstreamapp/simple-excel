/**
 * Summary, ratio and gate helpers over the `cells` array that `run.mjs` collects. Pure functions —
 * `run.mjs` owns the file I/O.
 */

export const BASELINE_ENGINE = 'sheetjs';

/**
 * Acceptance gates for a replacement engine. Ratio gates compare against the SheetJS cell with the
 * same op/dataset/size; absolute gates only look at the candidate.
 */
export const GATES = [
  {
    id: 'write-memory-100k',
    description: 'write RSS footprint <= 0.25x sheetjs (mixed 100k)',
    op: 'write',
    dataset: 'mixed',
    size: '100k',
    kind: 'ratio',
    metric: 'memory',
    maxRatio: 0.25,
  },
  {
    id: 'write-time-100k',
    description: 'write median time <= 1.0x sheetjs (mixed 100k)',
    op: 'write',
    dataset: 'mixed',
    size: '100k',
    kind: 'ratio',
    metric: 'time',
    maxRatio: 1.0,
  },
  {
    id: 'read-typed-time-100k',
    description: 'read-typed median time <= 1.0x sheetjs (mixed 100k)',
    op: 'read-typed',
    dataset: 'mixed',
    size: '100k',
    kind: 'ratio',
    metric: 'time',
    maxRatio: 1.0,
  },
  {
    id: 'read-typed-memory-100k',
    description: 'read-typed RSS footprint <= 0.5x sheetjs (mixed 100k)',
    op: 'read-typed',
    dataset: 'mixed',
    size: '100k',
    kind: 'ratio',
    metric: 'memory',
    maxRatio: 0.5,
  },
  {
    id: 'write-1m-succeeds',
    description: 'write 1M x 20 (mixed 1m) succeeds',
    op: 'write',
    dataset: 'mixed',
    size: '1m',
    kind: 'succeeds',
  },
  {
    id: 'write-18m-cells-succeeds',
    description: 'write 900k x 20 (mixed 18m-cells) succeeds',
    op: 'write',
    dataset: 'mixed',
    size: '18m-cells',
    kind: 'succeeds',
  },
  {
    id: 'first-byte-100k',
    description: 'first byte reaches the sink in < 100 ms (mixed 100k write)',
    op: 'write',
    dataset: 'mixed',
    size: '100k',
    kind: 'first-byte',
    maxMs: 100,
  },
];

export function median(values) {
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

export const formatMs = ms => (ms == null ? '—' : ms >= 1000 ? `${(ms / 1000).toFixed(2)} s` : `${ms.toFixed(ms >= 100 ? 0 : 1)} ms`);
export const formatMB = mb => (mb == null ? '—' : `${Math.round(mb).toLocaleString('en-US')} MB`);
export const formatBytes = bytes =>
  bytes == null ? '—' : bytes >= 1024 * 1024 ? `${(bytes / (1024 * 1024)).toFixed(1)} MB` : `${Math.round(bytes / 1024)} KB`;
export const formatRatio = ratio => (ratio == null ? '—' : `${ratio.toFixed(2)}x`);
export const formatInt = value => (value == null ? '—' : Number(value).toLocaleString('en-US'));

/** Memory metric = RSS footprint (peak minus the pre-warm-up baseline); falls back to the post-warm-up delta for old results. */
const metricValue = (cell, metric) =>
  metric === 'time' ? cell?.timing?.medianMs : (cell?.memory?.footprintMB ?? cell?.memory?.peakRssDeltaMB);

export function findCell(cells, { engine, op, dataset, size }) {
  return cells.find(cell => cell.engine === engine && cell.op === op && cell.dataset === dataset && cell.size === size);
}

/** Candidate ÷ baseline for a metric; null when either side is missing or failed. */
export function ratioVsBaseline(cells, cell, metric) {
  if (cell?.status !== 'ok' || cell.engine === BASELINE_ENGINE) {
    return null;
  }
  const baseline = findCell(cells, { ...cell, engine: BASELINE_ENGINE });
  const numerator = metricValue(cell, metric);
  const denominator = metricValue(baseline, metric);
  if (baseline?.status !== 'ok' || numerator == null || !denominator) {
    return null;
  }
  return numerator / denominator;
}

export function describeStatus(cell) {
  switch (cell.status) {
    case 'ok':
      return 'ok';
    case 'error': {
      const message = `${cell.error?.name ?? 'Error'}: ${cell.error?.message ?? ''}`.replace(/\s+/g, ' ').trim();
      return `ERROR (${cell.phase ?? 'run'}) ${message.length > 90 ? `${message.slice(0, 90)}…` : message}`;
    }
    case 'timeout':
      return `TIMEOUT (${formatMs(cell.timeoutMs)})`;
    case 'skipped':
      return `skipped: ${cell.reason ?? ''}`;
    case 'fixture-failed':
      return `fixture failed: ${cell.reason ?? ''}`;
    default:
      return cell.status;
  }
}

const SIZE_ORDER = ['1k', '10k', '100k', '1m', '18m-cells', 'wide-100k'];
const sizeRank = size => (SIZE_ORDER.includes(size) ? SIZE_ORDER.indexOf(size) : SIZE_ORDER.length);

function sortCells(cells, engineOrder) {
  return [...cells].sort(
    (a, b) =>
      sizeRank(a.size) - sizeRank(b.size) ||
      engineOrder.indexOf(a.engine) - engineOrder.indexOf(b.engine) ||
      a.engine.localeCompare(b.engine),
  );
}

function renderTable(header, rows) {
  const line = values => `| ${values.join(' | ')} |`;
  return [line(header), line(header.map(() => '---')), ...rows.map(line)].join('\n');
}

function renderWriteTable(cells, allCells, engineOrder) {
  const header = [
    'size',
    'engine',
    'status',
    'median',
    'min',
    'max',
    'time vs sheetjs',
    'RSS footprint',
    'mem vs sheetjs',
    'post-warm-up delta',
    'peak RSS',
    'output',
    'first byte',
  ];
  const rows = sortCells(cells, engineOrder).map(cell => [
    cell.size,
    cell.engine,
    describeStatus(cell),
    formatMs(cell.timing?.medianMs),
    formatMs(cell.timing?.minMs),
    formatMs(cell.timing?.maxMs),
    formatRatio(ratioVsBaseline(allCells, cell, 'time')),
    formatMB(metricValue(cell, 'memory')),
    formatRatio(ratioVsBaseline(allCells, cell, 'memory')),
    formatMB(cell.memory?.peakRssDeltaMB),
    formatMB(cell.memory?.peakRssMB),
    formatBytes(cell.bytes),
    formatMs(cell.firstByteMs),
  ]);
  return renderTable(header, rows);
}

function renderReadTable(cells, allCells, engineOrder) {
  const header = [
    'size',
    'engine',
    'status',
    'median',
    'min',
    'max',
    'time vs sheetjs',
    'RSS footprint',
    'mem vs sheetjs',
    'post-warm-up delta',
    'peak RSS',
    'rows read',
    'cells read',
  ];
  const rows = sortCells(cells, engineOrder).map(cell => [
    cell.size,
    cell.engine,
    describeStatus(cell),
    formatMs(cell.timing?.medianMs),
    formatMs(cell.timing?.minMs),
    formatMs(cell.timing?.maxMs),
    formatRatio(ratioVsBaseline(allCells, cell, 'time')),
    formatMB(metricValue(cell, 'memory')),
    formatRatio(ratioVsBaseline(allCells, cell, 'memory')),
    formatMB(cell.memory?.peakRssDeltaMB),
    formatMB(cell.memory?.peakRssMB),
    formatInt(cell.rowsRead),
    formatInt(cell.cellsRead),
  ]);
  return renderTable(header, rows);
}

/** Evaluates every gate for every non-baseline engine present in `cells`. */
export function evaluateGates(cells) {
  const engines = [...new Set(cells.map(cell => cell.engine))].filter(
    engine => engine !== BASELINE_ENGINE && cells.some(cell => cell.engine === engine && cell.status !== 'skipped'),
  );
  return engines.map(engine => {
    const gates = GATES.map(gate => {
      const cell = findCell(cells, { engine, op: gate.op, dataset: gate.dataset, size: gate.size });
      if (!cell) {
        return { ...gate, status: 'not-run', detail: 'cell not in this run' };
      }
      if (cell.status !== 'ok') {
        return { ...gate, status: cell.status === 'skipped' ? 'not-run' : 'fail', detail: describeStatus(cell) };
      }
      if (gate.kind === 'succeeds') {
        return {
          ...gate,
          status: 'pass',
          detail: `${formatMs(cell.timing.medianMs)}, ${formatMB(metricValue(cell, 'memory'))} RSS footprint`,
        };
      }
      if (gate.kind === 'first-byte') {
        if (cell.firstByteMs == null) {
          return { ...gate, status: 'fail', detail: 'engine reported no first-byte latency (not streaming)' };
        }
        return {
          ...gate,
          status: cell.firstByteMs < gate.maxMs ? 'pass' : 'fail',
          detail: `${formatMs(cell.firstByteMs)} (limit ${gate.maxMs} ms)`,
        };
      }
      const ratio = ratioVsBaseline(cells, cell, gate.metric);
      if (ratio == null) {
        const baseline = findCell(cells, { engine: BASELINE_ENGINE, op: gate.op, dataset: gate.dataset, size: gate.size });
        return {
          ...gate,
          status: 'not-run',
          detail: baseline ? `sheetjs baseline ${describeStatus(baseline)}` : 'sheetjs baseline not in this run',
        };
      }
      const candidateValue = metricValue(cell, gate.metric);
      const baselineValue = candidateValue / ratio;
      const format = gate.metric === 'time' ? formatMs : formatMB;
      return {
        ...gate,
        status: ratio <= gate.maxRatio ? 'pass' : 'fail',
        detail: `${formatRatio(ratio)} (${format(candidateValue)} vs ${format(baselineValue)}; limit ${gate.maxRatio}x)`,
      };
    });
    const verdict = gates.some(gate => gate.status === 'fail')
      ? 'FAIL'
      : gates.some(gate => gate.status === 'not-run')
        ? 'INCOMPLETE'
        : 'PASS';
    return { engine, verdict, gates };
  });
}

export function renderSummaryMarkdown(results) {
  const { cells, machine, engines, options, label } = results;
  const engineOrder = Object.keys(engines);
  const lines = [];
  lines.push(`# xlsx engine benchmark — ${label}`, '');
  lines.push(`Generated ${results.createdAt}. Results folder: \`${results.resultsDir}\`.`, '');
  lines.push('## Machine', '');
  lines.push(
    `- ${machine.cpuModel} (${machine.cpus} cores), ${machine.totalMemGB} GB RAM, ${machine.platform} ${machine.release} ${machine.arch}`,
  );
  lines.push(`- Node ${machine.nodeVersion} (V8 ${machine.v8}), child flags: \`${options.nodeFlags.join(' ')}\``);
  lines.push(
    `- Load average at start ${machine.loadAvgStart.map(value => value.toFixed(2)).join(' / ')}, at end ${(results.loadAvgEnd ?? []).map(value => value.toFixed(2)).join(' / ') || 'n/a'} (other processes may have been running)`,
  );
  lines.push(
    `- Options: runs=${options.runs}, warmup=${options.warmup}, source=${options.source}, timeout=${options.timeoutSeconds}s, seed=${options.seed}`,
    '',
  );
  lines.push('## Engines', '');
  lines.push(
    renderTable(
      ['engine', 'version', 'streaming', 'note'],
      engineOrder.map(name => [
        name,
        engines[name].version ?? '—',
        engines[name].supportsStreaming ? 'yes' : 'no',
        engines[name].skipReason ?? '',
      ]),
    ),
    '',
  );
  lines.push('## Metrics', '');
  lines.push('- `median`/`min`/`max`: wall time of the timed runs (after one warm-up run of the same cell).');
  lines.push(
    '- `RSS footprint`: highest RSS observed while the op ran (25 ms poller, post-run read, process high-water mark) minus the RSS before the warm-up (engine loaded, input data resident, gc() done). This is the ratio/gate metric.',
  );
  lines.push(
    '- `post-warm-up delta`: the same peak minus the RSS after warm-up + gc(). Undercounts engines whose warm-up footprint stays resident (V8 rarely returns pages), kept for reference.',
  );
  lines.push('- `peak RSS`: the absolute high-water RSS of the child process (input data included).');
  lines.push('- `first byte`: ms from the start of `write()` until the sink receives its first chunk (streaming writers only).');
  lines.push('- Ratios are candidate ÷ sheetjs for the same op/dataset/size; lower is better.', '');

  const ops = [...new Set(cells.map(cell => cell.op))];
  for (const op of ops) {
    const opCells = cells.filter(cell => cell.op === op);
    const datasets = [...new Set(opCells.map(cell => cell.dataset))];
    for (const dataset of datasets) {
      const subset = opCells.filter(cell => cell.dataset === dataset);
      lines.push(`## ${op} — ${dataset}`, '');
      const fixtures = [...new Set(subset.map(cell => cell.fixture?.producer).filter(Boolean))];
      if (fixtures.length > 0) {
        lines.push(`Fixtures written by: ${fixtures.join(', ')} (see \`fixture\` in results.json per cell).`, '');
      }
      lines.push(op === 'write' ? renderWriteTable(subset, cells, engineOrder) : renderReadTable(subset, cells, engineOrder), '');
    }
  }

  lines.push('## Gates', '');
  const evaluations = evaluateGates(cells);
  if (evaluations.length === 0) {
    lines.push('No candidate engines in this run (gates compare non-sheetjs engines against sheetjs).', '');
  }
  for (const { engine, verdict, gates } of evaluations) {
    lines.push(`### ${engine}: ${verdict}`, '');
    lines.push(
      renderTable(
        ['gate', 'status', 'detail'],
        gates.map(gate => [gate.description, gate.status.toUpperCase(), gate.detail]),
      ),
      '',
    );
  }

  const failures = cells.filter(cell => cell.status === 'error' || cell.status === 'timeout');
  if (failures.length > 0) {
    lines.push('## Failures', '');
    for (const cell of failures) {
      lines.push(`- **${cell.engine} / ${cell.op} / ${cell.dataset} / ${cell.size}** — ${describeStatus(cell)}`);
      if (cell.error?.stack) {
        lines.push('', '  ```', ...cell.error.stack.split('\n').map(line => `  ${line}`), '  ```', '');
      }
    }
    lines.push('');
  }
  return lines.join('\n');
}
