import { describe, expect, it } from 'vitest';
import { allFixtures, readFixture, type Fixture } from '../../../test/helpers/fixtures';
import type { SniffResult } from '../../types';
import { SNIFF_BYTES, sniff } from '../sniff';

function bytes(...values: readonly number[]): Uint8Array {
  return new Uint8Array(values);
}

function ascii(text: string): Uint8Array {
  return new TextEncoder().encode(text);
}

function withPrefix(prefix: readonly number[], text: string): Uint8Array {
  const tail = ascii(text);
  const combined = new Uint8Array(prefix.length + tail.length);
  combined.set(prefix, 0);
  combined.set(tail, prefix.length);
  return combined;
}

/** A CFB header followed by `name` written the way a directory entry stores it: UTF-16LE. */
function cfbWithStreamName(name: string): Uint8Array {
  const header = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  const utf16 = new Uint8Array(name.length * 2);
  for (let i = 0; i < name.length; i++) {
    utf16[i * 2] = name.charCodeAt(i) & 0xff;
  }
  const container = new Uint8Array(1024);
  container.set(header, 0);
  container.set(utf16, 512);
  return container;
}

describe('sniff', () => {
  it('recognizes a zip archive without deciding which kind it is', () => {
    // Every xlsx, xlsb and ods is a zip; only the content types tell them apart, so the facade decides that.
    expect(sniff(bytes(0x50, 0x4b, 0x03, 0x04, 0x14, 0x00))).toBe('zip');
  });

  it('reports an archive with no entries as empty', () => {
    expect(sniff(bytes(0x50, 0x4b, 0x05, 0x06, 0, 0, 0, 0))).toBe('empty');
  });

  it('EC-INPUT-ENCRYPTED: a CFB container with an EncryptedPackage stream is encrypted', () => {
    expect(sniff(cfbWithStreamName('EncryptedPackage'))).toBe('cfb-encrypted');
  });

  it('EC-INPUT-XLS-RENAMED: a CFB container without one is a legacy workbook', () => {
    expect(sniff(cfbWithStreamName('Workbook'))).toBe('cfb-legacy');
  });

  it('does not mistake an ASCII EncryptedPackage for the UTF-16LE stream name', () => {
    const container = new Uint8Array(1024);
    container.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0);
    container.set(ascii('EncryptedPackage'), 512);
    expect(sniff(container)).toBe('cfb-legacy');
  });

  it('EC-INPUT-SPREADSHEETML-2003: an XML document is xml', () => {
    expect(sniff(ascii('<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"/>'))).toBe('xml');
    expect(sniff(ascii('\n\t <?xml version="1.0"?><Workbook/>'))).toBe('xml');
    expect(sniff(withPrefix([0xef, 0xbb, 0xbf], '<?xml version="1.0"?><Workbook/>'))).toBe('xml');
  });

  it('recognizes HTML exports case-insensitively', () => {
    expect(sniff(ascii('<html><body><table></table></body></html>'))).toBe('html');
    expect(sniff(ascii('<!DOCTYPE html><html></html>'))).toBe('html');
    expect(sniff(ascii('  \r\n<HTML>'))).toBe('html');
    expect(sniff(ascii('<TABLE border="1"><tr><td>1</td></tr></TABLE>'))).toBe('html');
    expect(sniff(withPrefix([0xef, 0xbb, 0xbf], '<html>'))).toBe('html');
  });

  it('EC-INPUT-CSV-BYTES: csv and other plain text is text', () => {
    expect(sniff(ascii('"Id","Name","IsActive"\r\n"01t","Widget","true"\r\n'))).toBe('text');
    expect(sniff(ascii('Id,Name\n1,Zoë Ångström\n2,田中 太郎\n'))).toBe('text');
    expect(sniff(withPrefix([0xef, 0xbb, 0xbf], 'Id,Name\n1,a\n'))).toBe('text');
    // Mostly non-ASCII but valid UTF-8 is still text.
    expect(sniff(ascii('名前,住所\n田中 太郎,東京都\n佐藤 花子,大阪府\n'))).toBe('text');
  });

  it('EC-SNIFF-LEGACY-TEXT: a Windows-1252 csv with accented letters is text', () => {
    // `Name;City\nCafé Müller;Zürich\n` in Windows-1252: é is 0xE9 and ü is 0xFC, which is not UTF-8.
    const windows1252 = new Uint8Array([
      78, 97, 109, 101, 59, 67, 105, 116, 121, 10, 67, 97, 102, 233, 32, 77, 252, 108, 108, 101, 114, 59, 90, 252, 114, 105, 99, 104, 10,
    ]);
    expect(sniff(windows1252)).toBe('text');
    // Windows-1252's printable 0x80-0x9F range (smart quotes, the euro sign) is text too.
    expect(sniff(bytes(0x93, 0x51, 0x94, 0x2c, 0x80, 0x31, 0x30, 0x0d, 0x0a))).toBe('text');
  });

  it('EC-SNIFF-LEGACY-TEXT: a Shift-JIS csv is text', () => {
    // `名前,住所\r\n田中 太郎,東京都\r\n` in Shift-JIS.
    const shiftJis = new Uint8Array([
      150, 188, 145, 79, 44, 143, 90, 143, 138, 13, 10, 147, 99, 146, 134, 32, 145, 190, 152, 89, 44, 147, 140, 139, 158, 147, 115, 13, 10,
    ]);
    expect(sniff(shiftJis)).toBe('text');
  });

  it('EC-SNIFF-LEGACY-TEXT: UTF-16 with a byte-order mark is text, or xml/html when it says so', () => {
    const utf16 = (text: string, littleEndian: boolean): Uint8Array => {
      const out = new Uint8Array(2 + text.length * 2);
      out.set(littleEndian ? [0xff, 0xfe] : [0xfe, 0xff], 0);
      for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        out[2 + i * 2] = littleEndian ? code & 0xff : code >>> 8;
        out[3 + i * 2] = littleEndian ? code >>> 8 : code & 0xff;
      }
      return out;
    };
    // Excel's "Unicode Text" save: tab-separated UTF-16LE.
    expect(sniff(utf16('Id\tName\r\n1\tZoë Ångström 😀\r\n2\t田中 太郎\r\n', true))).toBe('text');
    expect(sniff(utf16('Id,Name\n1,Zoë\n', false))).toBe('text');
    expect(sniff(utf16(' <?xml version="1.0"?><Workbook/>', true))).toBe('xml');
    expect(sniff(utf16('<HTML><table></table></HTML>', false))).toBe('html');
    expect(sniff(bytes(0xff, 0xfe))).toBe('empty');
    // Long fields in a non-Latin script are almost no ASCII, and still text.
    expect(sniff(utf16(`${'東京都千代田区丸の内一丁目'.repeat(40)}\n`, true))).toBe('text');
    expect(sniff(utf16(`名前,住所\n${'北京市海淀区中关村大街'.repeat(40)}\n`, false))).toBe('text');
    // A byte-order mark in front of binary does not make it text when the units say binary: control characters,
    // unpaired surrogates or private-use code points.
    const binary = new Uint8Array(512);
    binary.set([0xff, 0xfe], 0);
    for (let i = 2; i < binary.length; i++) {
      binary[i] = i % 2 === 0 ? (i * 7) % 32 : 0;
    }
    expect(sniff(binary)).toBe('unknown');
    const privateUse = utf16('\uE000\uE123\uF8FF'.repeat(40), true);
    expect(sniff(privateUse)).toBe('unknown');
    const random = new Uint8Array(512);
    random.set([0xfe, 0xff], 0);
    for (let i = 2; i < random.length; i++) {
      random[i] = (i * 151 + 7) % 256;
    }
    expect(sniff(random)).toBe('unknown');
  });

  it('EC-SNIFF-LEGACY-TEXT: binary with no NUL byte but many control bytes stays unknown', () => {
    const noNul = new Uint8Array(512);
    for (let i = 0; i < noNul.length; i++) {
      noNul[i] = ((i * 97) % 255) + 1;
    }
    expect(noNul.includes(0)).toBe(false);
    expect(sniff(noNul)).toBe('unknown');
    // High bytes with a sprinkling of control bytes past the tolerance are binary too.
    const sparse = new Uint8Array(200).fill(0xe9);
    for (let i = 0; i < sparse.length; i += 20) {
      sparse[i] = 0x01;
    }
    expect(sniff(sparse)).toBe('unknown');
  });

  it('reports no bytes at all as empty', () => {
    expect(sniff(new Uint8Array(0))).toBe('empty');
    expect(sniff(bytes(0xef, 0xbb, 0xbf))).toBe('empty');
  });

  it('reports binary as unknown', () => {
    const random = new Uint8Array(512);
    for (let i = 0; i < random.length; i++) {
      random[i] = (i * 97) % 256;
    }
    expect(sniff(random)).toBe('unknown');
    // UTF-16LE without a byte-order mark: half the bytes are NUL and nothing says which encoding it is.
    expect(sniff(bytes(0x49, 0x00, 0x64, 0x00, 0x2c, 0x00, 0x4e, 0x00, 0x61, 0x00))).toBe('unknown');
    // A PDF is printable at the front and binary right after.
    const pdf = new Uint8Array(512);
    pdf.set(ascii('%PDF-1.7\n'), 0);
    pdf.fill(0x00, 9);
    expect(sniff(pdf)).toBe('unknown');
  });

  it('classifies from the bytes it has, without reading past them', () => {
    // Two printable letters, not the start of a signature it is allowed to assume.
    expect(sniff(bytes(0x50, 0x4b))).toBe('text');
    expect(sniff(bytes(0x50, 0x4b, 0x03))).toBe('unknown');
    expect(sniff(bytes(0xd0, 0xcf, 0x11))).toBe('unknown');
    expect(sniff(ascii('<'))).toBe('text');
  });
});

// ---- the corpus ----------------------------------------------------------------------------------------------------

/**
 * What each fixture must sniff as, from the catalog id the manifest tags it with. Everything else in the corpus is a
 * zip: the renamed `.ods` and `.xlsb` files included, because only their contents give them away.
 */
function expectedSniff(fixture: Fixture): SniffResult {
  if (fixture.tags.includes('EC-INPUT-XLS-RENAMED')) {
    return 'cfb-legacy';
  }
  if (fixture.tags.includes('EC-INPUT-ENCRYPTED')) {
    return 'cfb-encrypted';
  }
  if (fixture.tags.includes('EC-INPUT-SPREADSHEETML-2003')) {
    return 'xml';
  }
  if (fixture.tags.includes('EC-INPUT-CSV-BYTES')) {
    return 'text';
  }
  return 'zip';
}

describe('every fixture sniffs as its catalog entry says', () => {
  it.each(allFixtures().map(fixture => [fixture.id, fixture] as const))('%s', (_id, fixture) => {
    expect(sniff(readFixture(fixture).subarray(0, SNIFF_BYTES))).toBe(expectedSniff(fixture));
  });
});
