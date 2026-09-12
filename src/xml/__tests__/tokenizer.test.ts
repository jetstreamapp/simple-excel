import { describe, expect, it } from 'vitest';
import { isXlsxError } from '../../errors';
import { XmlTokenizer, type XmlTokenizerLimits } from '../tokenizer';

interface CollectOptions {
  /** Attribute local names read inside every `start` callback, so the events record what `attr` returned. */
  readonly probes?: readonly string[];
  readonly limits?: XmlTokenizerLimits;
}

const DEFAULT_PROBES = ['r', 't', 's', 'space'] as const;

/**
 * Feed `chunks` through a tokenizer and return one string per callback, which makes event sequences comparable
 * across different chunk splits with a single `toEqual`.
 */
function collect(chunks: readonly string[], options: CollectOptions = {}): string[] {
  const probes = options.probes ?? DEFAULT_PROBES;
  const events: string[] = [];
  const tokenizer: XmlTokenizer = new XmlTokenizer(
    {
      start(name: string): void {
        let event = `start:${name}`;
        for (const probe of probes) {
          const value = tokenizer.attr(probe);
          if (value !== undefined) {
            event += ` ${probe}=${JSON.stringify(value)}`;
          }
        }
        events.push(`${event} depth=${tokenizer.depth}`);
      },
      text(text: string): void {
        events.push(`text:${JSON.stringify(text)}`);
      },
      end(name: string): void {
        events.push(`end:${name}`);
      },
    },
    options.limits,
  );
  for (const chunk of chunks) {
    tokenizer.push(chunk);
  }
  tokenizer.end();
  return events;
}

function tokenize(xml: string, options: CollectOptions = {}): string[] {
  return collect([xml], options);
}

function splitAt(xml: string, ...indexes: readonly number[]): string[] {
  const bounds = [0, ...indexes, xml.length];
  const chunks: string[] = [];
  for (let i = 1; i < bounds.length; i++) {
    chunks.push(xml.slice(bounds[i - 1], bounds[i]));
  }
  return chunks;
}

const WORKSHEET_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
  ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
  `<dimension ref="A1:B20"/><sheetFormatPr defaultRowHeight="15"/>` +
  `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>0.30000000000000004</v></c></row></sheetData>` +
  `</worksheet>`;

describe('XmlTokenizer', () => {
  it('reads a worksheet part, skipping the XML declaration', () => {
    expect(tokenize(WORKSHEET_XML)).toEqual([
      'start:worksheet depth=1',
      'start:dimension depth=2',
      'end:dimension',
      'start:sheetFormatPr depth=2',
      'end:sheetFormatPr',
      'start:sheetData depth=2',
      'start:row r="1" depth=3',
      'start:c r="A1" t="s" depth=4',
      'start:v depth=5',
      'text:"0"',
      'end:v',
      'end:c',
      'start:c r="B1" depth=4',
      'start:v depth=5',
      'text:"0.30000000000000004"',
      'end:v',
      'end:c',
      'end:row',
      'end:sheetData',
      'end:worksheet',
    ]);
  });

  it('reads a shared string table with rich-text runs (EC-RICH-TEXT-RUNS)', () => {
    const xml =
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="215" uniqueCount="82">` +
      `<si><t>Object Api Name</t></si>` +
      `<si><t xml:space="preserve">  padded  </t></si>` +
      `<si><r><rPr><rFont val="Arial"/><b/></rPr><t xml:space="preserve">1. </t></r>` +
      `<r><t>Create one worksheet</t></r></si>` +
      `</sst>`;
    const events = tokenize(xml);
    expect(events.filter(event => event.startsWith('text:'))).toEqual([
      'text:"Object Api Name"',
      'text:"  padded  "',
      'text:"1. "',
      'text:"Create one worksheet"',
    ]);
    expect(events).toContain('start:t space="preserve" depth=3');
  });

  it('reads workbook.xml and its relationships by local attribute name', () => {
    const workbook =
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
      ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<sheets><sheet name="Probe" sheetId="1" r:id="rId1"/></sheets></workbook>`;
    expect(tokenize(workbook, { probes: ['name', 'sheetId', 'id'] })).toEqual([
      'start:workbook depth=1',
      'start:sheets depth=2',
      'start:sheet name="Probe" sheetId="1" id="rId1" depth=3',
      'end:sheet',
      'end:sheets',
      'end:workbook',
    ]);

    const rels =
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"` +
      ` Target="worksheets/sheet1.xml"/></Relationships>`;
    expect(tokenize(rels, { probes: ['Id', 'Target'] })).toEqual([
      'start:Relationships depth=1',
      'start:Relationship Id="rId1" Target="worksheets/sheet1.xml" depth=2',
      'end:Relationship',
      'end:Relationships',
    ]);
  });

  it('reads styles.xml', () => {
    const xml =
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;上午/下午 &quot;hh:mm"/></numFmts>` +
      `<fonts count="1"><font><name val="Calibri"/><family val="2"/><color theme="1"/><sz val="11"/></font></fonts>` +
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
      `<dxfs count="0"/></styleSheet>`;
    const events = tokenize(xml, { probes: ['numFmtId', 'formatCode', 'val'] });
    expect(events).toContain('start:numFmt numFmtId="164" formatCode="\\"上午/下午 \\"hh:mm" depth=3');
    expect(events).toContain('start:sz val="11" depth=4');
    expect(events.filter(event => event.startsWith('text:'))).toEqual([]);
  });

  it('strips namespace prefixes (EC-PREFIXED-ELEMENTS)', () => {
    const xml =
      `<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<x:sheetData><x:row x:r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>Id</x:t></x:is></x:c></x:row></x:sheetData>` +
      `</x:worksheet>`;
    expect(tokenize(xml)).toEqual([
      'start:worksheet depth=1',
      'start:sheetData depth=2',
      'start:row r="1" depth=3',
      'start:c r="A1" t="inlineStr" depth=4',
      'start:is depth=5',
      'start:t depth=6',
      'text:"Id"',
      'end:t',
      'end:is',
      'end:c',
      'end:row',
      'end:sheetData',
      'end:worksheet',
    ]);
  });

  it('reads Strict namespaces the same way (EC-STRICT-NAMESPACES)', () => {
    const xml =
      `<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">` +
      `<sheetData><row r="2"><c r="E2" t="d" s="3"><v>2024-02-29T12:00:00</v></c></row></sheetData></worksheet>`;
    expect(tokenize(xml)).toEqual([
      'start:worksheet depth=1',
      'start:sheetData depth=2',
      'start:row r="2" depth=3',
      'start:c r="E2" t="d" s="3" depth=4',
      'start:v depth=5',
      'text:"2024-02-29T12:00:00"',
      'end:v',
      'end:c',
      'end:row',
      'end:sheetData',
      'end:worksheet',
    ]);
  });

  it('delivers CDATA as text (EC-INLINE-STRINGS-CDATA)', () => {
    expect(tokenize('<is><t><![CDATA[x<y]]></t></is>')).toEqual(['start:is depth=1', 'start:t depth=2', 'text:"x<y"', 'end:t', 'end:is']);
    // CDATA joins the character data around it into one run, and entities inside it stay literal.
    expect(tokenize('<t>a<![CDATA[<b>&amp;]]>c</t>')).toEqual(['start:t depth=1', 'text:"a<b>&amp;c"', 'end:t']);
    expect(tokenize('<t><![CDATA[]]></t>')).toEqual(['start:t depth=1', 'end:t']);
  });

  it('decodes the predefined entities and numeric character references', () => {
    expect(tokenize('<t>&amp;&lt;&gt;&quot;&apos;</t>')).toEqual(['start:t depth=1', 'text:"&<>\\"\'"', 'end:t']);
    expect(tokenize('<t>a&#13;b&#x1F600;c&#65;</t>')).toEqual(['start:t depth=1', `text:${JSON.stringify('a\rb😀cA')}`, 'end:t']);
    expect(tokenize('<t>&amp;lt;</t>')).toEqual(['start:t depth=1', 'text:"&lt;"', 'end:t']);
  });

  it('EC-CRLF-NORMALIZED: raw line ends normalize, character references do not', () => {
    // XML 1.0 2.11 runs before parsing, so a raw CRLF or a lone CR is one LF wherever it appears; a carriage return
    // the producer meant to keep arrives as `&#13;` (or `_x000D_`, decoded a layer above) and survives.
    expect(tokenize('<t>a\r\nb\rc\nd</t>')).toEqual(['start:t depth=1', `text:${JSON.stringify('a\nb\nc\nd')}`, 'end:t']);
    expect(tokenize('<t>a&#13;\r\nb</t>')).toEqual(['start:t depth=1', `text:${JSON.stringify('a\r\nb')}`, 'end:t']);
    // A CRLF split across two pushes is still one line end.
    expect(collect(['<t>a\r', '\nb</t>'])).toEqual(['start:t depth=1', `text:${JSON.stringify('a\nb')}`, 'end:t']);
    expect(collect(['<t>a\r', 'b</t>'])).toEqual(['start:t depth=1', `text:${JSON.stringify('a\nb')}`, 'end:t']);
  });

  it('leaves references it does not recognize verbatim', () => {
    expect(tokenize('<t>&xxe; &foo &#zz; &#; 100% &</t>')).toEqual([
      'start:t depth=1',
      `text:${JSON.stringify('&xxe; &foo &#zz; &#; 100% &')}`,
      'end:t',
    ]);
  });

  it('treats a stray > in text as text and a quoted > as part of the attribute', () => {
    expect(tokenize('<t>a > b</t>')).toEqual(['start:t depth=1', 'text:"a > b"', 'end:t']);
    expect(tokenize('<c r="A>1"><v>1</v></c>', { probes: ['r'] })).toEqual([
      'start:c r="A>1" depth=1',
      'start:v depth=2',
      'text:"1"',
      'end:v',
      'end:c',
    ]);
  });

  it('skips comments and processing instructions', () => {
    expect(tokenize('<a><!-- a <comment> with --- dashes --><?ignore me?><b/></a>')).toEqual([
      'start:a depth=1',
      'start:b depth=2',
      'end:b',
      'end:a',
    ]);
  });

  it('reports self-closing elements as a start and an end', () => {
    expect(tokenize('<row r="21" ht="15.75" customHeight="1"/>')).toEqual(['start:row r="21" depth=1', 'end:row']);
    expect(tokenize('<c r="C2" s="8" />')).toEqual(['start:c r="C2" s="8" depth=1', 'end:c']);
  });

  it('reads attributes quoted either way and returns undefined for the rest', () => {
    const events: string[] = [];
    const tokenizer: XmlTokenizer = new XmlTokenizer({
      start(name: string): void {
        events.push(`${name} space=${String(tokenizer.attr('space'))} missing=${String(tokenizer.attr('missing'))}`);
      },
      text(): void {},
      end(): void {},
    });
    tokenizer.push(`<t xml:space='preserve'>  x  </t>`);
    tokenizer.end();
    expect(events).toEqual(['t space=preserve missing=undefined']);
    // Outside a start callback there is no element to read from.
    expect(tokenizer.attr('space')).toBeUndefined();
  });

  it('does not expose namespace declarations as attributes', () => {
    // Stripping prefixes would otherwise make `xmlns:r` look like the row/cell reference attribute `r`.
    const events = tokenize(WORKSHEET_XML);
    expect(events[0]).toBe('start:worksheet depth=1');
    expect(tokenize('<a xmlns="urn:x" xmlns:r="urn:y" r:id="rId1"/>', { probes: ['id', 'xmlns'] })).toEqual([
      'start:a id="rId1" depth=1',
      'end:a',
    ]);
  });

  it('strips a leading byte order mark', () => {
    expect(tokenize('\ufeff<t>x</t>')).toEqual(['start:t depth=1', 'text:"x"', 'end:t']);
  });

  it('delivers whitespace-only text between elements', () => {
    expect(tokenize('<a>\n  <b/>\n</a>')).toEqual(['start:a depth=1', 'text:"\\n  "', 'start:b depth=2', 'end:b', 'text:"\\n"', 'end:a']);
  });
});

describe('XmlTokenizer chunk boundaries', () => {
  const FUZZ_DOCUMENT = buildFuzzDocument();

  function buildFuzzDocument(): string {
    let xml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<!-- generated by the fuzz test --><dimension ref="A1:D30"/>` +
      `<cols><col customWidth="1" min="1" max="1" width="2.88"/></cols>` +
      `<sheetData>`;
    for (let row = 1; row <= 24; row++) {
      xml +=
        `<x:row r="${row}" spans="1:4"><x:c r="A${row}" t="s"><x:v>${row}</x:v></x:c>` +
        `<c r="B${row}"><v>0.3000000000000${row}</v></c>` +
        `<c r="C${row}" t="inlineStr"><is><t xml:space="preserve"> a &amp; b &#13;&#x1F600; </t></is></c>` +
        `<c r="D${row}" t="inlineStr"><is><t><![CDATA[x<y & z]]></t></is></c>` +
        `<c r="E${row}" s="8"/></x:row>`;
    }
    return `${xml}</sheetData><?calc later?><mergeCells count="1"><mergeCell ref="B9:B24"/></mergeCells></x:worksheet>`;
  }

  it('produces the same events however the input is split', () => {
    expect(FUZZ_DOCUMENT.length).toBeGreaterThan(2000);
    const expected = tokenize(FUZZ_DOCUMENT);
    for (let index = 0; index <= FUZZ_DOCUMENT.length; index++) {
      expect(collect(splitAt(FUZZ_DOCUMENT, index)), `split at ${index}`).toEqual(expected);
    }
  });

  it('produces the same events for random split triples', () => {
    const expected = tokenize(FUZZ_DOCUMENT);
    let state = 12345;
    const nextIndex = (): number => {
      state = (Math.imul(state, 1103515245) + 12345) >>> 0;
      return state % (FUZZ_DOCUMENT.length + 1);
    };
    for (let round = 0; round < 500; round++) {
      const indexes = [nextIndex(), nextIndex(), nextIndex()].toSorted((left, right) => left - right);
      expect(collect(splitAt(FUZZ_DOCUMENT, ...indexes)), `split at ${indexes.join(',')}`).toEqual(expected);
    }
  });

  it('joins an entity split across chunks', () => {
    expect(collect(['<t>a&am', 'p;b</t>'])).toEqual(['start:t depth=1', 'text:"a&b"', 'end:t']);
    expect(collect(['<t>a&#x1F6', '00;b</t>'])).toEqual(['start:t depth=1', `text:${JSON.stringify('a😀b')}`, 'end:t']);
    expect(collect(['<t>a&', '', 'amp', ';b</t>'])).toEqual(['start:t depth=1', 'text:"a&b"', 'end:t']);
  });

  it('joins a CDATA marker split across chunks', () => {
    expect(collect(['<t><![CDA', 'TA[x<y]', ']', '>z</t>'])).toEqual(['start:t depth=1', 'text:"x<yz"', 'end:t']);
  });

  it('delivers one text run per character-data run, never one per chunk', () => {
    const chunks = ['<t>abc', 'def', 'ghi</t>'];
    expect(collect(chunks)).toEqual(['start:t depth=1', 'text:"abcdefghi"', 'end:t']);
  });
});

describe('XmlTokenizer limits and hostile input', () => {
  /** The `XlsxError` code a tokenizer run throws, so every expectation reads as one `toBe` on a classified code. */
  function thrownCode(run: () => void): string {
    try {
      run();
    } catch (error) {
      return isXlsxError(error) ? error.code : `unclassified: ${String(error)}`;
    }
    return 'nothing thrown';
  }

  it('rejects DOCTYPE outright (EC-XXE-DOCTYPE)', () => {
    const payload =
      `<?xml version="1.0"?>` +
      `<!DOCTYPE sst [<!ENTITY xxe SYSTEM "file:///etc/passwd"><!ENTITY lol "lololol">]>` +
      `<sst><si><t>&xxe;</t></si></sst>`;
    expect(thrownCode(() => tokenize(payload))).toBe('XML_DOCTYPE');
    expect(thrownCode(() => collect(['<sst><!D', 'OCTYPE sst><si/></sst>']))).toBe('XML_DOCTYPE');
    expect(thrownCode(() => tokenize('<sst><!ENTITY lol "lol"><si/></sst>'))).toBe('XML_DOCTYPE');
  });

  it('caps nesting depth (EC-XML-DEEP-NESTING)', () => {
    const deep = `<si>${'<r>'.repeat(300)}deep${'</r>'.repeat(300)}</si>`;
    expect(thrownCode(() => tokenize(deep))).toBe('LIMIT_EXCEEDED');
    expect(() => tokenize(`<si>${'<r>'.repeat(100)}ok${'</r>'.repeat(100)}</si>`)).not.toThrow();
    expect(thrownCode(() => tokenize('<a><b><c>x</c></b></a>', { limits: { maxDepth: 2 } }))).toBe('LIMIT_EXCEEDED');
  });

  it('caps the length of a single text run', () => {
    const limits: XmlTokenizerLimits = { maxTextLength: 1024 };
    expect(thrownCode(() => tokenize(`<t>${'x'.repeat(2048)}</t>`, { limits }))).toBe('LIMIT_EXCEEDED');
    // The cap covers a run assembled from several chunks, not just one chunk's worth.
    expect(thrownCode(() => collect(['<t>', 'x'.repeat(600), 'x'.repeat(600), '</t>'], { limits }))).toBe('LIMIT_EXCEEDED');
  });

  it('rejects structurally broken markup', () => {
    const broken = [
      '<a><b></a></b>',
      '<a><b>text</b>',
      '<a>text',
      '<a></b></a>',
      '<a></a></a>',
      '<a><!-- unterminated',
      '<a><![CDATA[unterminated</a>',
      '<a><?pi unterminated',
      '<>x</>',
      '<a><c r="unterminated></a>',
    ];
    for (const xml of broken) {
      expect(
        thrownCode(() => tokenize(xml)),
        xml,
      ).toBe('XML_MALFORMED');
    }
  });

  it('rejects a tag longer than a mebicharacter', () => {
    expect(thrownCode(() => tokenize(`<c ${'a'.repeat(1024 * 1024 + 8)}="1"/>`))).toBe('XML_MALFORMED');
  });
});

describe('XmlTokenizer throughput', () => {
  it('tokenizes a synthetic 50 MB worksheet', () => {
    let block = '';
    for (let row = 0; row < 4000; row++) {
      // No colon anywhere in the block: that is the worst case for resolving local names.
      block +=
        `<row r="${row + 1}">` +
        `<c r="A${row + 1}" t="s"><v>${row % 997}</v></c>` +
        `<c r="B${row + 1}"><v>0.30000000000000004</v></c>` +
        `<c r="C${row + 1}" t="inlineStr"><is><t>Object Api Name ${row}</t></is></c>` +
        `<c r="D${row + 1}" s="3"><v>45352.5</v></c>` +
        `<c r="E${row + 1}"/>` +
        `</row>`;
    }
    const blockCount = Math.ceil((50 * 1024 * 1024) / block.length);
    const totalCharacters = block.length * blockCount + 64;

    let cellCount = 0;
    let textCount = 0;
    const tokenizer: XmlTokenizer = new XmlTokenizer({
      start(name: string): void {
        if (name === 'c') {
          cellCount++;
        }
      },
      text(): void {
        textCount++;
      },
      end(): void {},
    });

    const startedAt = performance.now();
    tokenizer.push('<worksheet><sheetData>');
    for (let i = 0; i < blockCount; i++) {
      tokenizer.push(block);
    }
    tokenizer.push('</sheetData></worksheet>');
    tokenizer.end();
    const elapsedSeconds = (performance.now() - startedAt) / 1000;

    const megabytes = totalCharacters / (1024 * 1024);
    process.stdout.write(
      `\n  XmlTokenizer: ${(megabytes / elapsedSeconds).toFixed(1)} MB/s ` +
        `(${megabytes.toFixed(1)} MB, ${cellCount.toLocaleString('en-US')} cells, ${elapsedSeconds.toFixed(2)}s)\n`,
    );
    expect(cellCount).toBe(5 * 4000 * blockCount);
    expect(textCount).toBe(4 * 4000 * blockCount);
  }, 120_000);
});
